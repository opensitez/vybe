use super::{PascalParser, Rule};
use pest::Parser;
use pest::iterators::Pair;
use vybe_ast::*;

const PASCAL_HELPER_TARGET_PREFIX: &str = "__pascal_helper_target__:";
const PASCAL_VARIANT_FIELD_MARKER: &str = "__pascal_variant_field__";

pub fn parse(source: &str) -> Result<Module, String> {
    let source = source.trim_start_matches('\u{feff}');
    let pairs =
        PascalParser::parse(Rule::program, source).map_err(|e| format!("Parse error: {}", e))?;
    let mut body = Vec::new();
    let mut imports = Vec::new();
    let mut name = "main".to_string();
    let mut is_unit = false;

    for pair in pairs {
        if pair.as_rule() != Rule::program {
            continue;
        }
        for inner in pair.into_inner() {
            match inner.as_rule() {
                Rule::program_heading => {
                    // program Foo; or unit Foo;
                    is_unit = inner
                        .as_str()
                        .trim_start()
                        .get(..4)
                        .is_some_and(|prefix| prefix.eq_ignore_ascii_case("unit"));
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
                        } else if p.as_rule() == Rule::uses_item {
                            let mut unit_name: Option<String> = None;
                            let mut source_path: Option<String> = None;
                            for part in p.into_inner() {
                                match part.as_rule() {
                                    Rule::identifier => {
                                        if unit_name.is_none() {
                                            unit_name = Some(part.as_str().to_string());
                                        }
                                    }
                                    Rule::string_literal => {
                                        let raw = part.as_str();
                                        let path = raw
                                            .strip_prefix('\'')
                                            .and_then(|s| s.strip_suffix('\''))
                                            .unwrap_or(raw)
                                            .replace("''", "'");
                                        source_path = Some(path);
                                    }
                                    _ => {}
                                }
                            }
                            if let Some(path) = source_path.or(unit_name) {
                                imports.push(Import {
                                    kind: ImportKind::Simple { path, alias: None },
                                    span,
                                });
                            }
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
    lower_pascal_gotos_in_body(&mut body);
    lower_pascal_file_io(&mut body);
    normalize_pascal_free_function_overloads(&mut body);

    // Synthesize minimal RTL classes at the top of every Pascal program
    // before constructor-call rewriting so known runtime base types exist
    // during later passes.
    if !is_unit {
        body.insert(0, synthesize_tinterfacedobject_class());
        body.insert(0, synthesize_exception_class());
    }

    let uses_gcl = imports.iter().any(|import| match &import.kind {
        ImportKind::Simple { path, .. }
        | ImportKind::Named { path, .. }
        | ImportKind::Wildcard { path, .. }
        | ImportKind::Default { path, .. } => vybe_platform_plib::emitter::gcl::is_gcl_unit(path),
    });
    if uses_gcl {
        normalize_pascal_gcl_form_classes(&mut body);
        normalize_pascal_gcl_exprs(&mut body);
    }
    lower_pascal_operator_overloads(&mut body);

    // Now that class declarations are stable, rewrite `TFoo.Create(args)` (Pascal's
    // constructor invocation syntax) into the canonical `New { class: TFoo, args }`
    // AST so every language ends up with the same instantiation node.
    let mut class_names: std::collections::HashSet<String> = body
        .iter()
        .filter_map(|s| match &s.kind {
            StmtKind::ClassDecl { name, .. } | StmtKind::StructDecl { name, .. } => {
                Some(name.to_lowercase())
            }
            _ => None,
        })
        .collect();
    if uses_gcl {
        for class in vybe_platform_plib::emitter::gcl::gcl_classes() {
            class_names.insert(class.name.to_lowercase());
        }
    }
    let mut class_display_names: Vec<(String, String)> = body
        .iter()
        .filter_map(|s| match &s.kind {
            StmtKind::ClassDecl { name, .. } | StmtKind::StructDecl { name, .. } => {
                Some((name.to_lowercase(), name.clone()))
            }
            _ => None,
        })
        .collect();
    if uses_gcl {
        for class in vybe_platform_plib::emitter::gcl::gcl_classes() {
            class_display_names.push((class.name.to_lowercase(), class.name.to_string()));
        }
    }
    let (static_methods, static_values) = collect_static_members(&body);
    let static_var_params = collect_static_var_param_indices(&body);
    for stmt in body.iter_mut() {
        rewrite_constructor_calls_stmt(stmt, &class_names, &static_methods);
    }
    for stmt in body.iter_mut() {
        rewrite_pascal_rtti_stmt(stmt, &class_display_names);
    }

    let static_properties = collect_pascal_static_properties(&body);
    if !static_properties.is_empty() {
        rewrite_pascal_static_properties(&mut body, &static_properties);
    }
    for stmt in body.iter_mut() {
        rewrite_static_method_calls_stmt(stmt, &static_methods, &static_values);
    }
    for stmt in body.iter_mut() {
        mark_static_var_args_stmt(stmt, &static_var_params);
    }
    lower_pascal_method_pointers(&mut body);
    let zero_arg_instance_methods = collect_zero_arg_instance_methods(&body);
    for stmt in body.iter_mut() {
        rewrite_zero_arg_instance_method_refs_stmt(stmt, &zero_arg_instance_methods);
    }
    rewrite_bare_parameterless_method_refs(&mut body);
    let indexed_properties = collect_pascal_indexed_properties(&body);
    if !indexed_properties.is_empty() {
        rewrite_pascal_indexed_properties(&mut body, &indexed_properties);
    }
    let polymorphic_class_names = collect_pascal_polymorphic_class_names(&body);
    for stmt in body.iter_mut() {
        erase_pascal_class_value_type_hints_stmt(stmt, &polymorphic_class_names);
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
    let record_array_types = collect_record_array_types(&body, &struct_names);
    let early_enum_type_counts = collect_enum_type_counts(&body);
    normalize_pascal_enum_indexed_array_decls(&mut body, &early_enum_type_counts);
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
    lower_pascal_array_value_semantics(&mut body);
    default_init_const_bounded_arrays(&mut body);
    rewrite_pascal_fixed_array_bounds(&mut body);
    for stmt in body.iter_mut() {
        materialize_record_array_setlength_stmt(stmt, &record_array_types);
    }
    let mut string_vars = std::collections::HashSet::new();
    let mut zero_based_loop_vars = std::collections::HashSet::new();
    for stmt in body.iter_mut() {
        rewrite_zero_based_string_indexes_stmt(stmt, &mut string_vars, &mut zero_based_loop_vars);
    }
    rewrite_pascal_datetime_arithmetic(&mut body);
    let enum_type_names = collect_enum_type_names(&body);
    let enum_type_counts = collect_enum_type_counts(&body);
    let enum_member_ordinals = collect_enum_member_ordinals(&body);
    default_init_enum_indexed_arrays(&mut body, &enum_type_counts);
    rewrite_pascal_enum_ordinals(&mut body, &enum_member_ordinals);
    rename_shadowing_pascal_set_vars(&mut body, &enum_member_ordinals);
    rewrite_pascal_set_semantics(&mut body, &enum_type_names);
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
    if uses_gcl {
        normalize_pascal_gcl_exprs(&mut body);
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
                        default_init_struct_locals_stmt(
                            s,
                            struct_names,
                            explicit_ctor_record_names,
                        );
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

#[derive(Clone)]
struct PascalOverloadCandidate {
    internal_name: String,
    params: Vec<Param>,
    return_type: Option<String>,
    order: usize,
}

fn normalize_pascal_free_function_overloads(body: &mut Vec<Statement>) {
    let mut grouped: std::collections::BTreeMap<String, Vec<PascalOverloadCandidate>> =
        std::collections::BTreeMap::new();
    let enum_members = collect_enum_member_types(body);

    for (order, stmt) in body.iter().enumerate() {
        let StmtKind::FunctionDecl {
            name,
            params,
            return_type,
            ..
        } = &stmt.kind
        else {
            continue;
        };
        if name.contains('.') {
            continue;
        }
        grouped
            .entry(name.to_lowercase())
            .or_default()
            .push(PascalOverloadCandidate {
                internal_name: name.clone(),
                params: params.clone(),
                return_type: return_type.clone(),
                order,
            });
    }

    grouped.retain(|_, candidates| candidates.len() > 1);
    if grouped.is_empty() {
        return;
    }

    let mut rename_by_order = std::collections::HashMap::new();
    let mut return_types = std::collections::HashMap::new();
    for (lowered, candidates) in grouped.iter_mut() {
        for (idx, candidate) in candidates.iter_mut().enumerate() {
            candidate.internal_name = format!("__pascal_overload_{}_{}", lowered, idx);
            rename_by_order.insert(candidate.order, candidate.internal_name.clone());
            return_types.insert(
                candidate.internal_name.to_lowercase(),
                candidate.return_type.clone(),
            );
        }
    }

    for (order, stmt) in body.iter_mut().enumerate() {
        if let Some(new_name) = rename_by_order.get(&order) {
            if let StmtKind::FunctionDecl { name, .. } = &mut stmt.kind {
                *name = new_name.clone();
            }
        }
    }

    let all_candidates: std::collections::HashMap<String, Vec<PascalOverloadCandidate>> =
        grouped.into_iter().collect();
    let mut scope = std::collections::HashMap::new();
    for stmt in body.iter_mut() {
        rewrite_pascal_overload_stmt(stmt, &all_candidates, &return_types, &enum_members, &mut scope);
    }
}

fn collect_enum_member_types(body: &[Statement]) -> std::collections::HashMap<String, String> {
    fn visit_stmt(stmt: &Statement, out: &mut std::collections::HashMap<String, String>) {
        match &stmt.kind {
            StmtKind::EnumDecl { name, members, .. } => {
                for member in members {
                    out.insert(member.name.to_lowercase(), name.to_lowercase());
                }
            }
            StmtKind::ClassDecl { members, .. } | StmtKind::StructDecl { members, .. } => {
                for member in members {
                    match member {
                        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
                            visit_stmt(stmt, out)
                        }
                        ClassMember::Constructor { body, .. } => {
                            for stmt in body {
                                visit_stmt(stmt, out);
                            }
                        }
                        _ => {}
                    }
                }
            }
            StmtKind::Block(body) | StmtKind::FunctionDecl { body, .. } => {
                for stmt in body {
                    visit_stmt(stmt, out);
                }
            }
            _ => {}
        }
    }

    let mut out = std::collections::HashMap::new();
    for stmt in body {
        visit_stmt(stmt, &mut out);
    }
    out
}

fn rewrite_pascal_overload_stmt(
    stmt: &mut Statement,
    overloads: &std::collections::HashMap<String, Vec<PascalOverloadCandidate>>,
    return_types: &std::collections::HashMap<String, Option<String>>,
    enum_members: &std::collections::HashMap<String, String>,
    scope: &mut std::collections::HashMap<String, String>,
) {
    match &mut stmt.kind {
        StmtKind::VarDecl { declarations, kind } => {
            for decl in declarations {
                if let (BindingPattern::Ident(name), Some(type_hint)) =
                    (&decl.pattern, &decl.type_hint)
                {
                    scope.insert(name.to_lowercase(), bare_type_name(type_hint).to_lowercase());
                }
                if *kind == VarDeclKind::Const {
                    if let BindingPattern::Ident(name) = &decl.pattern {
                        if let Some(init) = &decl.init {
                            if let Some(type_name) = pascal_overload_expr_type(
                                init,
                                return_types,
                                enum_members,
                                scope,
                            ) {
                                scope.insert(name.to_lowercase(), type_name);
                            }
                        }
                    }
                }
                if let Some(init) = &mut decl.init {
                    rewrite_pascal_overload_expr(init, overloads, return_types, enum_members, scope);
                }
            }
        }
        StmtKind::FunctionDecl { params, body, .. } => {
            let mut scoped = scope.clone();
            for param in params {
                if let Some(type_hint) = &param.type_hint {
                    scoped.insert(param.name.to_lowercase(), bare_type_name(type_hint).to_lowercase());
                }
            }
            for stmt in body {
                rewrite_pascal_overload_stmt(stmt, overloads, return_types, enum_members, &mut scoped);
            }
        }
        StmtKind::Block(body) => {
            let mut scoped = scope.clone();
            for stmt in body {
                rewrite_pascal_overload_stmt(stmt, overloads, return_types, enum_members, &mut scoped);
            }
        }
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) | StmtKind::Throw { expr: Some(expr), .. } => {
            rewrite_pascal_overload_expr(expr, overloads, return_types, enum_members, scope);
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_pascal_overload_expr(target, overloads, return_types, enum_members, scope);
            }
            rewrite_pascal_overload_expr(value, overloads, return_types, enum_members, scope);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_pascal_overload_expr(target, overloads, return_types, enum_members, scope);
            rewrite_pascal_overload_expr(value, overloads, return_types, enum_members, scope);
        }
        StmtKind::If { cond, then_body, elifs, else_body } => {
            rewrite_pascal_overload_expr(cond, overloads, return_types, enum_members, scope);
            let mut then_scope = scope.clone();
            for stmt in then_body {
                rewrite_pascal_overload_stmt(stmt, overloads, return_types, enum_members, &mut then_scope);
            }
            for (cond, body) in elifs {
                rewrite_pascal_overload_expr(cond, overloads, return_types, enum_members, scope);
                let mut scoped = scope.clone();
                for stmt in body {
                    rewrite_pascal_overload_stmt(stmt, overloads, return_types, enum_members, &mut scoped);
                }
            }
            if let Some(body) = else_body {
                let mut scoped = scope.clone();
                for stmt in body {
                    rewrite_pascal_overload_stmt(stmt, overloads, return_types, enum_members, &mut scoped);
                }
            }
        }
        StmtKind::While { cond, body, .. } | StmtKind::DoWhile { cond, body, .. } => {
            rewrite_pascal_overload_expr(cond, overloads, return_types, enum_members, scope);
            let mut scoped = scope.clone();
            for stmt in body {
                rewrite_pascal_overload_stmt(stmt, overloads, return_types, enum_members, &mut scoped);
            }
        }
        StmtKind::For { init, cond, update, body } => {
            let mut scoped = scope.clone();
            if let Some(init) = init {
                rewrite_pascal_overload_stmt(init, overloads, return_types, enum_members, &mut scoped);
            }
            if let Some(cond) = cond {
                rewrite_pascal_overload_expr(cond, overloads, return_types, enum_members, &mut scoped);
            }
            if let Some(update) = update {
                rewrite_pascal_overload_expr(update, overloads, return_types, enum_members, &mut scoped);
            }
            for stmt in body {
                rewrite_pascal_overload_stmt(stmt, overloads, return_types, enum_members, &mut scoped);
            }
        }
        StmtKind::ForIn { iter, body, .. } => {
            rewrite_pascal_overload_expr(iter, overloads, return_types, enum_members, scope);
            let mut scoped = scope.clone();
            for stmt in body {
                rewrite_pascal_overload_stmt(stmt, overloads, return_types, enum_members, &mut scoped);
            }
        }
        StmtKind::Switch { expr, cases, default } => {
            rewrite_pascal_overload_expr(expr, overloads, return_types, enum_members, scope);
            for case in cases {
                for cond in &mut case.conditions {
                    match cond {
                        CaseCondition::Value(expr) | CaseCondition::Comparison { expr, .. } => {
                            rewrite_pascal_overload_expr(expr, overloads, return_types, enum_members, scope);
                        }
                        CaseCondition::Range { from, to } => {
                            rewrite_pascal_overload_expr(from, overloads, return_types, enum_members, scope);
                            rewrite_pascal_overload_expr(to, overloads, return_types, enum_members, scope);
                        }
                    }
                }
                let mut scoped = scope.clone();
                for stmt in &mut case.body {
                    rewrite_pascal_overload_stmt(stmt, overloads, return_types, enum_members, &mut scoped);
                }
            }
            if let Some(body) = default {
                let mut scoped = scope.clone();
                for stmt in body {
                    rewrite_pascal_overload_stmt(stmt, overloads, return_types, enum_members, &mut scoped);
                }
            }
        }
        StmtKind::ClassDecl { members, .. } | StmtKind::StructDecl { members, .. } => {
            for member in members {
                match member {
                    ClassMember::Field { init: Some(init), type_hint, name, .. } => {
                        if let Some(type_hint) = type_hint {
                            scope.insert(name.to_lowercase(), bare_type_name(type_hint).to_lowercase());
                        }
                        rewrite_pascal_overload_expr(init, overloads, return_types, enum_members, scope);
                    }
                    ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
                        let mut scoped = scope.clone();
                        rewrite_pascal_overload_stmt(stmt, overloads, return_types, enum_members, &mut scoped);
                    }
                    ClassMember::Constructor { params, body, .. } => {
                        let mut scoped = scope.clone();
                        for param in params {
                            if let Some(type_hint) = &param.type_hint {
                                scoped.insert(param.name.to_lowercase(), bare_type_name(type_hint).to_lowercase());
                            }
                        }
                        for stmt in body {
                            rewrite_pascal_overload_stmt(stmt, overloads, return_types, enum_members, &mut scoped);
                        }
                    }
                    _ => {}
                }
            }
        }
        _ => {}
    }
}

fn rewrite_pascal_overload_expr(
    expr: &mut Expression,
    overloads: &std::collections::HashMap<String, Vec<PascalOverloadCandidate>>,
    return_types: &std::collections::HashMap<String, Option<String>>,
    enum_members: &std::collections::HashMap<String, String>,
    scope: &std::collections::HashMap<String, String>,
) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            for arg in args.iter_mut() {
                rewrite_pascal_overload_expr(&mut arg.value, overloads, return_types, enum_members, scope);
            }
            if let ExprKind::Ident(name) = &mut callee.kind {
                if let Some(candidates) = overloads.get(&name.to_lowercase()) {
                    if let Some(best) = choose_pascal_overload(candidates, args, return_types, enum_members, scope) {
                        *name = best.internal_name.clone();
                    }
                }
            } else {
                rewrite_pascal_overload_expr(callee, overloads, return_types, enum_members, scope);
            }
        }
        ExprKind::Binary { left, right, .. } => {
            rewrite_pascal_overload_expr(left, overloads, return_types, enum_members, scope);
            rewrite_pascal_overload_expr(right, overloads, return_types, enum_members, scope);
        }
        ExprKind::Unary { expr, .. } | ExprKind::RefLoad(expr) => {
            rewrite_pascal_overload_expr(expr, overloads, return_types, enum_members, scope);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_pascal_overload_expr(cond, overloads, return_types, enum_members, scope);
            rewrite_pascal_overload_expr(then, overloads, return_types, enum_members, scope);
            rewrite_pascal_overload_expr(else_, overloads, return_types, enum_members, scope);
        }
        ExprKind::Member { object, .. } => {
            rewrite_pascal_overload_expr(object, overloads, return_types, enum_members, scope);
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_pascal_overload_expr(object, overloads, return_types, enum_members, scope);
            rewrite_pascal_overload_expr(index, overloads, return_types, enum_members, scope);
        }
        ExprKind::New { class, args } => {
            rewrite_pascal_overload_expr(class, overloads, return_types, enum_members, scope);
            for arg in args {
                rewrite_pascal_overload_expr(&mut arg.value, overloads, return_types, enum_members, scope);
            }
        }
        ExprKind::Assign { target, value } => {
            rewrite_pascal_overload_expr(target, overloads, return_types, enum_members, scope);
            rewrite_pascal_overload_expr(value, overloads, return_types, enum_members, scope);
        }
        ExprKind::Array(items) => {
            for item in items {
                if let Some(key) = &mut item.key {
                    rewrite_pascal_overload_expr(key, overloads, return_types, enum_members, scope);
                }
                rewrite_pascal_overload_expr(&mut item.value, overloads, return_types, enum_members, scope);
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) => {
            for item in items {
                rewrite_pascal_overload_expr(item, overloads, return_types, enum_members, scope);
            }
        }
        ExprKind::Cast { expr, .. } | ExprKind::IsType { expr, .. } => {
            rewrite_pascal_overload_expr(expr, overloads, return_types, enum_members, scope);
        }
        _ => {}
    }
}

fn choose_pascal_overload<'a>(
    candidates: &'a [PascalOverloadCandidate],
    args: &[Argument],
    return_types: &std::collections::HashMap<String, Option<String>>,
    enum_members: &std::collections::HashMap<String, String>,
    scope: &std::collections::HashMap<String, String>,
) -> Option<&'a PascalOverloadCandidate> {
    candidates
        .iter()
        .filter_map(|candidate| {
            let required = candidate.params.iter().filter(|p| p.default.is_none()).count();
            if args.len() < required || args.len() > candidate.params.len() {
                return None;
            }
            let mut score = if args.len() == candidate.params.len() { 0 } else { 10 };
            for (arg, param) in args.iter().zip(candidate.params.iter()) {
                score += pascal_overload_arg_score(
                    arg,
                    param,
                    return_types,
                    enum_members,
                    scope,
                )?;
            }
            Some((score, candidate.params.len(), candidate.order, candidate))
        })
        .min_by_key(|(score, params_len, order, _)| (*score, *params_len, *order))
        .map(|(_, _, _, candidate)| candidate)
}

fn pascal_overload_arg_score(
    arg: &Argument,
    param: &Param,
    return_types: &std::collections::HashMap<String, Option<String>>,
    enum_members: &std::collections::HashMap<String, String>,
    scope: &std::collections::HashMap<String, String>,
) -> Option<usize> {
    if matches!(param.pass_by, PassBy::Ref | PassBy::Out) && !pascal_overload_is_assignable(&arg.value) {
        return None;
    }
    let param_type = param
        .type_hint
        .as_deref()
        .map(|hint| bare_type_name(hint).to_lowercase());
    let arg_type = pascal_overload_expr_type(&arg.value, return_types, enum_members, scope);
    match (arg_type.as_deref(), param_type.as_deref()) {
        (_, None) => Some(20),
        (Some("nil"), Some("pointer")) => Some(0),
        (Some("nil"), _) => Some(12),
        (Some(arg), Some(param)) if arg == param => Some(0),
        (Some("char"), Some("string")) => Some(1),
        (Some("string"), Some("char")) => Some(50),
        (Some("integer"), Some("real" | "single" | "double" | "extended")) => Some(2),
        (Some("integer"), Some("longint" | "shortint" | "byte" | "word" | "cardinal" | "int64")) => Some(1),
        (Some(arg), Some("integer")) if enum_members.values().any(|ty| ty.eq_ignore_ascii_case(arg)) => Some(4),
        (Some(_), Some(_)) => Some(30),
        (None, Some(_)) => Some(15),
    }
}

fn pascal_overload_expr_type(
    expr: &Expression,
    return_types: &std::collections::HashMap<String, Option<String>>,
    enum_members: &std::collections::HashMap<String, String>,
    scope: &std::collections::HashMap<String, String>,
) -> Option<String> {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(_)) => Some("integer".to_string()),
        ExprKind::Lit(Literal::Float(_)) => Some("real".to_string()),
        ExprKind::Lit(Literal::Bool(_)) => Some("boolean".to_string()),
        ExprKind::Lit(Literal::Null) => Some("nil".to_string()),
        ExprKind::Lit(Literal::Char(_)) => Some("char".to_string()),
        ExprKind::Lit(Literal::Str(s)) if s.chars().count() == 1 => Some("char".to_string()),
        ExprKind::Lit(Literal::Str(_)) => Some("string".to_string()),
        ExprKind::Ident(name) => scope
            .get(&name.to_lowercase())
            .cloned()
            .or_else(|| enum_members.get(&name.to_lowercase()).cloned()),
        ExprKind::Call { callee, .. } => {
            if let ExprKind::Ident(name) = &callee.kind {
                return_types
                    .get(&name.to_lowercase())
                    .and_then(|ty| ty.clone())
                    .map(|ty| bare_type_name(&ty).to_lowercase())
            } else {
                None
            }
        }
        ExprKind::Cast { type_name, .. } => Some(bare_type_name(type_name).to_lowercase()),
        ExprKind::Unary { op: UnaryOp::Not, .. } => Some("boolean".to_string()),
        ExprKind::Binary { op, .. } => match op {
            BinOp::Eq | BinOp::NotEq | BinOp::Lt | BinOp::LtEq | BinOp::Gt | BinOp::GtEq => {
                Some("boolean".to_string())
            }
            _ => None,
        },
        _ => None,
    }
}

fn pascal_overload_is_assignable(expr: &Expression) -> bool {
    matches!(
        expr.kind,
        ExprKind::Ident(_) | ExprKind::Member { .. } | ExprKind::Index { .. } | ExprKind::RefLoad(_)
    )
}

#[derive(Clone)]
struct PascalOperatorMethod {
    name: String,
    params: Vec<Param>,
    return_type: Option<String>,
}

fn lower_pascal_operator_overloads(body: &mut Vec<Statement>) {
    let operators = collect_pascal_operator_methods(body);
    if operators.is_empty() {
        return;
    }
    let mut scope = std::collections::HashMap::new();
    lower_pascal_operator_overload_body(body, &operators, &mut scope);
}

fn collect_pascal_operator_methods(
    body: &[Statement],
) -> std::collections::HashMap<String, Vec<PascalOperatorMethod>> {
    let mut out = std::collections::HashMap::new();
    for stmt in body {
        let (StmtKind::ClassDecl { name, members, .. } | StmtKind::StructDecl { name, members, .. }) =
            &stmt.kind
        else {
            continue;
        };
        let methods = out.entry(name.to_lowercase()).or_insert_with(Vec::new);
        for member in members {
            let ClassMember::Method(method) = member else {
                continue;
            };
            let StmtKind::FunctionDecl {
                name,
                params,
                return_type,
                modifiers,
                ..
            } = &method.kind
            else {
                continue;
            };
            let operator_name = name.strip_prefix("operator_").unwrap_or(name);
            if modifiers.is_static && is_pascal_operator_method_name(operator_name) {
                methods.push(PascalOperatorMethod {
                    name: operator_name.to_string(),
                    params: params.clone(),
                    return_type: return_type.clone(),
                });
            }
        }
    }
    out
}

fn is_pascal_operator_method_name(name: &str) -> bool {
    matches!(
        name.to_ascii_lowercase().as_str(),
        "add"
            | "subtract"
            | "multiply"
            | "divide"
            | "intdivide"
            | "modulus"
            | "lessthan"
            | "lessthanorequal"
            | "greaterthan"
            | "greaterthanorequal"
            | "equal"
            | "notequal"
            | "bitwiseand"
            | "bitwiseor"
            | "bitwisexor"
            | "leftshift"
            | "rightshift"
            | "negative"
            | "not"
            | "inc"
            | "dec"
            | "implicit"
            | "explicit"
    )
}

fn lower_pascal_operator_overload_body(
    body: &mut [Statement],
    operators: &std::collections::HashMap<String, Vec<PascalOperatorMethod>>,
    scope: &mut std::collections::HashMap<String, String>,
) {
    for stmt in body {
        lower_pascal_operator_overload_stmt(stmt, operators, scope);
    }
}

fn lower_pascal_operator_overload_stmt(
    stmt: &mut Statement,
    operators: &std::collections::HashMap<String, Vec<PascalOperatorMethod>>,
    scope: &mut std::collections::HashMap<String, String>,
) {
    match &mut stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let (BindingPattern::Ident(name), Some(type_hint)) =
                    (&decl.pattern, &decl.type_hint)
                {
                    scope.insert(name.to_lowercase(), bare_type_name(type_hint).to_lowercase());
                }
                if let Some(init) = &mut decl.init {
                    lower_pascal_operator_overload_expr(init, operators, scope);
                    if let (BindingPattern::Ident(_), Some(type_hint)) =
                        (&decl.pattern, &decl.type_hint)
                    {
                        if let Some(converted) =
                            pascal_implicit_operator_call(bare_type_name(type_hint), init, operators, scope)
                        {
                            *init = converted;
                        }
                    }
                }
            }
        }
        StmtKind::FunctionDecl { params, body, .. } => {
            let mut scoped = scope.clone();
            for param in params {
                if let Some(type_hint) = &param.type_hint {
                    scoped.insert(param.name.to_lowercase(), bare_type_name(type_hint).to_lowercase());
                }
            }
            lower_pascal_operator_overload_body(body, operators, &mut scoped);
        }
        StmtKind::Block(body) => {
            let mut scoped = scope.clone();
            lower_pascal_operator_overload_body(body, operators, &mut scoped);
        }
        StmtKind::Expr(expr) => {
            if let Some(call) = pascal_inc_dec_operator_stmt_expr(expr, operators, scope) {
                *expr = call;
            } else {
                lower_pascal_operator_overload_expr(expr, operators, scope);
            }
        }
        StmtKind::Assign { targets, value } => {
            for target in targets.iter_mut() {
                lower_pascal_operator_overload_expr(target, operators, scope);
            }
            lower_pascal_operator_overload_expr(value, operators, scope);
            if targets.len() == 1 {
                if let Some(call) = pascal_inc_dec_operator_assignment(&targets[0], value, operators, scope) {
                    stmt.kind = StmtKind::Expr(call);
                    return;
                }
                if let Some(target_type) = pascal_operator_expr_type(&targets[0], scope, operators) {
                    if let Some(converted) =
                        pascal_implicit_operator_call(&target_type, value, operators, scope)
                    {
                        *value = converted;
                    }
                }
            }
        }
        StmtKind::CompoundAssign { target, value, op } => {
            lower_pascal_operator_overload_expr(target, operators, scope);
            lower_pascal_operator_overload_expr(value, operators, scope);
            if let Some(method) = pascal_operator_method_for_compound(*op) {
                if let Some(call) =
                    pascal_binary_operator_call(method, target, value, operators, scope)
                {
                    stmt.kind = StmtKind::Assign {
                        targets: vec![target.clone()],
                        value: call,
                    };
                }
            }
        }
        StmtKind::If { cond, then_body, elifs, else_body } => {
            lower_pascal_operator_overload_expr(cond, operators, scope);
            let mut then_scope = scope.clone();
            lower_pascal_operator_overload_body(then_body, operators, &mut then_scope);
            for (cond, body) in elifs {
                lower_pascal_operator_overload_expr(cond, operators, scope);
                let mut scoped = scope.clone();
                lower_pascal_operator_overload_body(body, operators, &mut scoped);
            }
            if let Some(body) = else_body {
                let mut scoped = scope.clone();
                lower_pascal_operator_overload_body(body, operators, &mut scoped);
            }
        }
        StmtKind::While { cond, body, .. } | StmtKind::DoWhile { cond, body, .. } => {
            lower_pascal_operator_overload_expr(cond, operators, scope);
            let mut scoped = scope.clone();
            lower_pascal_operator_overload_body(body, operators, &mut scoped);
        }
        StmtKind::For { init, cond, update, body } => {
            let mut scoped = scope.clone();
            if let Some(init) = init {
                lower_pascal_operator_overload_stmt(init, operators, &mut scoped);
            }
            if let Some(cond) = cond {
                lower_pascal_operator_overload_expr(cond, operators, &mut scoped);
            }
            if let Some(update) = update {
                lower_pascal_operator_overload_expr(update, operators, &mut scoped);
            }
            lower_pascal_operator_overload_body(body, operators, &mut scoped);
        }
        StmtKind::ForIn { iter, body, .. } => {
            lower_pascal_operator_overload_expr(iter, operators, scope);
            let mut scoped = scope.clone();
            lower_pascal_operator_overload_body(body, operators, &mut scoped);
        }
        StmtKind::Switch { expr, cases, default } => {
            lower_pascal_operator_overload_expr(expr, operators, scope);
            for case in cases {
                for cond in &mut case.conditions {
                    match cond {
                        CaseCondition::Value(expr) | CaseCondition::Comparison { expr, .. } => {
                            lower_pascal_operator_overload_expr(expr, operators, scope)
                        }
                        CaseCondition::Range { from, to } => {
                            lower_pascal_operator_overload_expr(from, operators, scope);
                            lower_pascal_operator_overload_expr(to, operators, scope);
                        }
                    }
                }
                let mut scoped = scope.clone();
                lower_pascal_operator_overload_body(&mut case.body, operators, &mut scoped);
            }
            if let Some(body) = default {
                let mut scoped = scope.clone();
                lower_pascal_operator_overload_body(body, operators, &mut scoped);
            }
        }
        StmtKind::ClassDecl { members, .. } | StmtKind::StructDecl { members, .. } => {
            for member in members {
                match member {
                    ClassMember::Field { name, type_hint, init: Some(init), .. } => {
                        if let Some(type_hint) = type_hint {
                            scope.insert(name.to_lowercase(), bare_type_name(type_hint).to_lowercase());
                        }
                        lower_pascal_operator_overload_expr(init, operators, scope);
                    }
                    ClassMember::Method(method) | ClassMember::NestedType(method) => {
                        let mut scoped = scope.clone();
                        lower_pascal_operator_overload_stmt(method, operators, &mut scoped);
                    }
                    ClassMember::Constructor { params, body, .. } => {
                        let mut scoped = scope.clone();
                        for param in params {
                            if let Some(type_hint) = &param.type_hint {
                                scoped.insert(param.name.to_lowercase(), bare_type_name(type_hint).to_lowercase());
                            }
                        }
                        lower_pascal_operator_overload_body(body, operators, &mut scoped);
                    }
                    _ => {}
                }
            }
        }
        StmtKind::Return(Some(expr)) | StmtKind::Throw { expr: Some(expr), .. } => {
            lower_pascal_operator_overload_expr(expr, operators, scope);
        }
        _ => {}
    }
}

fn lower_pascal_operator_overload_expr(
    expr: &mut Expression,
    operators: &std::collections::HashMap<String, Vec<PascalOperatorMethod>>,
    scope: &std::collections::HashMap<String, String>,
) {
    match &mut expr.kind {
        ExprKind::Binary { op, left, right } => {
            lower_pascal_operator_overload_expr(left, operators, scope);
            lower_pascal_operator_overload_expr(right, operators, scope);
            if let Some(method) = pascal_operator_method_for_binop(*op) {
                if let Some(call) = pascal_binary_operator_call(method, left, right, operators, scope) {
                    *expr = call;
                } else if *op == BinOp::Gt {
                    if let Some(call) =
                        pascal_binary_operator_call("LessThan", right, left, operators, scope)
                    {
                        *expr = call;
                    }
                }
            }
        }
        ExprKind::Unary { op, expr: inner } => {
            lower_pascal_operator_overload_expr(inner, operators, scope);
            let method = match op {
                UnaryOp::Neg => Some("Negative"),
                UnaryOp::Not => Some("Not"),
                _ => None,
            };
            if let Some(method) = method {
                if let Some(call) = pascal_unary_operator_call(method, inner, operators, scope) {
                    *expr = call;
                }
            }
        }
        ExprKind::Cast { expr: inner, type_name } => {
            lower_pascal_operator_overload_expr(inner, operators, scope);
            if let Some(call) = pascal_explicit_operator_call(type_name, inner, operators, scope) {
                *expr = call;
            }
        }
        ExprKind::Call { callee, args, .. } => {
            lower_pascal_operator_overload_expr(callee, operators, scope);
            for arg in args {
                lower_pascal_operator_overload_expr(&mut arg.value, operators, scope);
            }
        }
        ExprKind::Member { object, .. } => lower_pascal_operator_overload_expr(object, operators, scope),
        ExprKind::Index { object, index, .. } => {
            lower_pascal_operator_overload_expr(object, operators, scope);
            lower_pascal_operator_overload_expr(index, operators, scope);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            lower_pascal_operator_overload_expr(cond, operators, scope);
            lower_pascal_operator_overload_expr(then, operators, scope);
            lower_pascal_operator_overload_expr(else_, operators, scope);
        }
        ExprKind::Assign { target, value } => {
            lower_pascal_operator_overload_expr(target, operators, scope);
            lower_pascal_operator_overload_expr(value, operators, scope);
        }
        ExprKind::Array(items) => {
            for item in items {
                if let Some(key) = &mut item.key {
                    lower_pascal_operator_overload_expr(key, operators, scope);
                }
                lower_pascal_operator_overload_expr(&mut item.value, operators, scope);
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) => {
            for item in items {
                lower_pascal_operator_overload_expr(item, operators, scope);
            }
        }
        ExprKind::New { class, args } => {
            lower_pascal_operator_overload_expr(class, operators, scope);
            for arg in args {
                lower_pascal_operator_overload_expr(&mut arg.value, operators, scope);
            }
        }
        _ => {}
    }
}

fn pascal_operator_method_for_binop(op: BinOp) -> Option<&'static str> {
    match op {
        BinOp::Add => Some("Add"),
        BinOp::Sub => Some("Subtract"),
        BinOp::Mul => Some("Multiply"),
        BinOp::Div => Some("Divide"),
        BinOp::IDiv => Some("IntDivide"),
        BinOp::Mod => Some("Modulus"),
        BinOp::Lt => Some("LessThan"),
        BinOp::LtEq => Some("LessThanOrEqual"),
        BinOp::Gt => Some("GreaterThan"),
        BinOp::GtEq => Some("GreaterThanOrEqual"),
        BinOp::Eq => Some("Equal"),
        BinOp::NotEq => Some("NotEqual"),
        BinOp::And | BinOp::BitAnd => Some("BitwiseAnd"),
        BinOp::Or | BinOp::BitOr => Some("BitwiseOr"),
        BinOp::Xor | BinOp::BitXor => Some("BitwiseXor"),
        BinOp::Shl => Some("LeftShift"),
        BinOp::Shr => Some("RightShift"),
        _ => None,
    }
}

fn pascal_operator_method_for_compound(op: CompoundOp) -> Option<&'static str> {
    match op {
        CompoundOp::Add => Some("Add"),
        CompoundOp::Sub => Some("Subtract"),
        CompoundOp::Mul => Some("Multiply"),
        CompoundOp::Div => Some("Divide"),
        CompoundOp::IDiv => Some("IntDivide"),
        CompoundOp::Mod => Some("Modulus"),
        CompoundOp::BitAnd => Some("BitwiseAnd"),
        CompoundOp::BitOr => Some("BitwiseOr"),
        CompoundOp::BitXor => Some("BitwiseXor"),
        CompoundOp::Shl => Some("LeftShift"),
        CompoundOp::Shr => Some("RightShift"),
        _ => None,
    }
}

fn pascal_binary_operator_call(
    method: &str,
    left: &Expression,
    right: &Expression,
    operators: &std::collections::HashMap<String, Vec<PascalOperatorMethod>>,
    scope: &std::collections::HashMap<String, String>,
) -> Option<Expression> {
    let left_type = pascal_operator_expr_type(left, scope, operators)?;
    let right_type = pascal_operator_expr_type(right, scope, operators);
    let owner = choose_pascal_operator_owner(method, &[Some(left_type), right_type], operators)?;
    Some(pascal_static_operator_call(&owner, method, vec![left.clone(), right.clone()]))
}

fn pascal_unary_operator_call(
    method: &str,
    value: &Expression,
    operators: &std::collections::HashMap<String, Vec<PascalOperatorMethod>>,
    scope: &std::collections::HashMap<String, String>,
) -> Option<Expression> {
    let value_type = pascal_operator_expr_type(value, scope, operators)?;
    let owner = choose_pascal_operator_owner(method, &[Some(value_type)], operators)?;
    Some(pascal_static_operator_call(&owner, method, vec![value.clone()]))
}

fn pascal_explicit_operator_call(
    target_type: &str,
    value: &Expression,
    operators: &std::collections::HashMap<String, Vec<PascalOperatorMethod>>,
    scope: &std::collections::HashMap<String, String>,
) -> Option<Expression> {
    let value_type = pascal_operator_expr_type(value, scope, operators)?;
    let owner = choose_pascal_operator_owner("Explicit", &[Some(value_type)], operators)?;
    let method = find_pascal_operator(&owner, "Explicit", operators)?;
    if method
        .return_type
        .as_deref()
        .is_some_and(|ty| bare_type_name(ty).eq_ignore_ascii_case(target_type))
    {
        Some(pascal_static_operator_call(&owner, "Explicit", vec![value.clone()]))
    } else {
        None
    }
}

fn pascal_implicit_operator_call(
    target_type: &str,
    value: &Expression,
    operators: &std::collections::HashMap<String, Vec<PascalOperatorMethod>>,
    scope: &std::collections::HashMap<String, String>,
) -> Option<Expression> {
    let target = target_type.to_lowercase();
    if pascal_operator_expr_type(value, scope, operators)
        .is_some_and(|ty| ty.eq_ignore_ascii_case(&target))
    {
        return None;
    }
    let methods = operators.get(&target)?;
    let method = methods.iter().find(|method| {
        method.name.eq_ignore_ascii_case("Implicit")
            && method.params.len() == 1
            && pascal_operator_arg_matches(&method.params[0], value, scope)
    })?;
    Some(pascal_static_operator_call(&target, &method.name, vec![value.clone()]))
}

fn pascal_inc_dec_operator_stmt_expr(
    expr: &Expression,
    operators: &std::collections::HashMap<String, Vec<PascalOperatorMethod>>,
    scope: &std::collections::HashMap<String, String>,
) -> Option<Expression> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    let ExprKind::Ident(name) = &callee.kind else {
        return None;
    };
    let method = if name.eq_ignore_ascii_case("Inc") {
        "Inc"
    } else if name.eq_ignore_ascii_case("Dec") {
        "Dec"
    } else {
        return None;
    };
    let target = args.first()?.value.clone();
    let target_type = pascal_operator_expr_type(&target, scope, operators)?;
    if find_pascal_operator(&target_type, method, operators).is_none() {
        return None;
    }
    let mut call_args = vec![target];
    if let Some(arg) = args.get(1) {
        call_args.push(arg.value.clone());
    }
    Some(pascal_static_operator_call(&target_type, method, call_args))
}

fn pascal_inc_dec_operator_assignment(
    target: &Expression,
    value: &Expression,
    operators: &std::collections::HashMap<String, Vec<PascalOperatorMethod>>,
    scope: &std::collections::HashMap<String, String>,
) -> Option<Expression> {
    let ExprKind::Binary { op, left, right } = &value.kind else {
        return None;
    };
    let method = match op {
        BinOp::Add => "Inc",
        BinOp::Sub => "Dec",
        _ => return None,
    };
    if !pascal_expr_same_place(target, left) {
        return None;
    }
    let target_type = pascal_operator_expr_type(target, scope, operators)?;
    let operator = find_pascal_operator(&target_type, method, operators)?;
    let mut args = vec![target.clone()];
    if operator.params.len() > 1 || !matches!(right.kind, ExprKind::Lit(Literal::Int(1))) {
        args.push((**right).clone());
    }
    Some(pascal_static_operator_call(&target_type, method, args))
}

fn choose_pascal_operator_owner(
    method: &str,
    arg_types: &[Option<String>],
    operators: &std::collections::HashMap<String, Vec<PascalOperatorMethod>>,
) -> Option<String> {
    let first_type = arg_types.iter().flatten().find(|ty| operators.contains_key(*ty))?;
    let methods = operators.get(first_type)?;
    methods
        .iter()
        .find(|candidate| {
            candidate.name.eq_ignore_ascii_case(method)
                && candidate.params.len() == arg_types.len()
                && candidate
                    .params
                    .iter()
                    .zip(arg_types.iter())
                    .all(|(param, arg_type)| match (param.type_hint.as_deref(), arg_type) {
                        (Some(param_type), Some(arg_type)) => {
                            bare_type_name(param_type).eq_ignore_ascii_case(arg_type)
                                || is_pascal_numeric_type(param_type)
                        }
                        (Some(_), None) => true,
                        (None, _) => true,
                    })
        })
        .map(|_| first_type.clone())
}

fn find_pascal_operator<'a>(
    type_name: &str,
    method: &str,
    operators: &'a std::collections::HashMap<String, Vec<PascalOperatorMethod>>,
) -> Option<&'a PascalOperatorMethod> {
    operators
        .get(&type_name.to_lowercase())?
        .iter()
        .find(|candidate| candidate.name.eq_ignore_ascii_case(method))
}

fn pascal_static_operator_call(type_name: &str, method: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident(type_name)),
            field: method.to_string(),
            null_safe: false,
        })),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

fn pascal_operator_arg_matches(
    param: &Param,
    value: &Expression,
    scope: &std::collections::HashMap<String, String>,
) -> bool {
    let Some(param_type) = param.type_hint.as_deref() else {
        return true;
    };
    match pascal_operator_expr_type(value, scope, &std::collections::HashMap::new()) {
        Some(value_type) => {
            bare_type_name(param_type).eq_ignore_ascii_case(&value_type)
                || is_pascal_numeric_type(param_type)
        }
        None => true,
    }
}

fn pascal_operator_expr_type(
    expr: &Expression,
    scope: &std::collections::HashMap<String, String>,
    operators: &std::collections::HashMap<String, Vec<PascalOperatorMethod>>,
) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => scope.get(&name.to_lowercase()).cloned(),
        ExprKind::Member { object, field, .. } => {
            let object_type = pascal_operator_expr_type(object, scope, operators)?;
            Some(format!("{}.{}", object_type, field).to_lowercase())
        }
        ExprKind::Lit(Literal::Int(_)) => Some("integer".to_string()),
        ExprKind::Lit(Literal::Float(_)) => Some("real".to_string()),
        ExprKind::Lit(Literal::Bool(_)) => Some("boolean".to_string()),
        ExprKind::Lit(Literal::Char(_)) => Some("char".to_string()),
        ExprKind::Lit(Literal::Str(_)) => Some("string".to_string()),
        ExprKind::Call { callee, .. } => {
            let ExprKind::Member { object, field, .. } = &callee.kind else {
                return None;
            };
            let ExprKind::Ident(type_name) = &object.kind else {
                return None;
            };
            operators
                .get(&type_name.to_lowercase())
                .and_then(|methods| {
                    methods
                        .iter()
                        .find(|method| method.name.eq_ignore_ascii_case(field))
                        .and_then(|method| method.return_type.as_deref())
                })
                .map(|ty| bare_type_name(ty).to_lowercase())
                .or_else(|| Some(format!("{}.{}", type_name, field).to_lowercase()))
        }
        _ => None,
    }
}

fn pascal_expr_same_place(a: &Expression, b: &Expression) -> bool {
    match (&a.kind, &b.kind) {
        (ExprKind::Ident(a), ExprKind::Ident(b)) => a.eq_ignore_ascii_case(b),
        (
            ExprKind::Member { object: ao, field: af, .. },
            ExprKind::Member { object: bo, field: bf, .. },
        ) => af.eq_ignore_ascii_case(bf) && pascal_expr_same_place(ao, bo),
        (
            ExprKind::Index { object: ao, index: ai, .. },
            ExprKind::Index { object: bo, index: bi, .. },
        ) => pascal_expr_same_place(ao, bo) && format!("{:?}", ai.kind) == format!("{:?}", bi.kind),
        _ => false,
    }
}

fn is_pascal_numeric_type(type_name: &str) -> bool {
    matches!(
        bare_type_name(type_name).to_ascii_lowercase().as_str(),
        "integer"
            | "longint"
            | "shortint"
            | "byte"
            | "word"
            | "cardinal"
            | "int64"
            | "real"
            | "single"
            | "double"
            | "extended"
    )
}

struct PascalMethodPointerInfo {
    aliases: std::collections::HashMap<String, String>,
    fields: std::collections::HashMap<String, std::collections::HashMap<String, String>>,
    methods: std::collections::HashMap<String, std::collections::HashSet<String>>,
    parents: std::collections::HashMap<String, Vec<String>>,
    routines: std::collections::HashMap<String, Vec<Param>>,
}

fn lower_pascal_method_pointers(body: &mut Vec<Statement>) {
    let info = collect_pascal_method_pointer_info(body);
    let mut var_types = std::collections::HashMap::new();
    lower_pascal_method_pointer_body(body, &info, &mut var_types);
}

fn collect_pascal_method_pointer_info(body: &[Statement]) -> PascalMethodPointerInfo {
    let mut info = PascalMethodPointerInfo {
        aliases: std::collections::HashMap::new(),
        fields: std::collections::HashMap::new(),
        methods: std::collections::HashMap::new(),
        parents: std::collections::HashMap::new(),
        routines: std::collections::HashMap::new(),
    };
    for stmt in body {
        match &stmt.kind {
            StmtKind::VarDecl { declarations, kind } if *kind == VarDeclKind::Const => {
                for decl in declarations {
                    if let (BindingPattern::Ident(name), Some(type_hint)) = (&decl.pattern, &decl.type_hint) {
                        info.aliases.insert(name.to_lowercase(), type_hint.to_lowercase());
                    }
                }
            }
            StmtKind::FunctionDecl { name, params, .. } if !name.contains('.') => {
                info.routines.insert(name.to_lowercase(), params.clone());
            }
            StmtKind::ClassDecl { name, parents, members, .. } => {
                let class = name.to_lowercase();
                info.parents.insert(
                    class.clone(),
                    parents.iter().map(|p| bare_type_name(p).to_lowercase()).collect(),
                );
                collect_pascal_method_pointer_members(&mut info, class, members);
            }
            StmtKind::StructDecl { name, members, .. } => {
                collect_pascal_method_pointer_members(&mut info, name.to_lowercase(), members);
            }
            _ => {}
        }
    }
    info
}

fn collect_pascal_method_pointer_members(
    info: &mut PascalMethodPointerInfo,
    class: String,
    members: &[ClassMember],
) {
    let field_map = info.fields.entry(class.clone()).or_default();
    let method_set = info.methods.entry(class).or_default();
    for member in members {
        match member {
            ClassMember::Field { name, type_hint, .. } => {
                if let Some(type_hint) = type_hint {
                    field_map.insert(name.to_lowercase(), type_hint.to_lowercase());
                }
            }
            ClassMember::Method(method) => {
                if let StmtKind::FunctionDecl { name, params, .. } = &method.kind {
                    if params.is_empty() {
                        method_set.insert(name.to_lowercase());
                    }
                }
            }
            _ => {}
        }
    }
}

fn lower_pascal_method_pointer_body(
    body: &mut [Statement],
    info: &PascalMethodPointerInfo,
    var_types: &mut std::collections::HashMap<String, String>,
) {
    for stmt in body {
        lower_pascal_method_pointer_stmt(stmt, info, var_types);
    }
}

fn lower_pascal_method_pointer_stmt(
    stmt: &mut Statement,
    info: &PascalMethodPointerInfo,
    var_types: &mut std::collections::HashMap<String, String>,
) {
    match &mut stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations.iter_mut() {
                if let Some(init) = &mut decl.init {
                    lower_pascal_method_pointer_expr(init, info, var_types);
                }
            }
            for decl in declarations {
                if let (BindingPattern::Ident(name), Some(type_hint)) = (&decl.pattern, &decl.type_hint) {
                    var_types.insert(name.to_lowercase(), type_hint.to_lowercase());
                }
            }
        }
        StmtKind::Assign { targets, value } if targets.len() == 1 => {
            lower_pascal_method_pointer_expr(value, info, var_types);
            if pascal_method_pointer_target_type(&targets[0], info, var_types)
                .is_some_and(|ty| is_pascal_procedural_type(&ty, info))
            {
                if let Some(lambda) = pascal_method_pointer_lambda(value, info, var_types) {
                    *value = lambda;
                }
            }
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                lower_pascal_method_pointer_expr(target, info, var_types);
            }
            lower_pascal_method_pointer_expr(value, info, var_types);
        }
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            lower_pascal_method_pointer_expr(expr, info, var_types);
        }
        StmtKind::FunctionDecl { params, body, .. } => {
            let mut scoped = var_types.clone();
            for param in params {
                if let Some(type_hint) = &param.type_hint {
                    scoped.insert(param.name.to_lowercase(), type_hint.to_lowercase());
                }
            }
            lower_pascal_method_pointer_body(body, info, &mut scoped);
        }
        StmtKind::ClassDecl { members, .. } | StmtKind::StructDecl { members, .. } => {
            for member in members {
                match member {
                    ClassMember::Method(method) | ClassMember::NestedType(method) => {
                        let mut scoped = std::collections::HashMap::new();
                        lower_pascal_method_pointer_stmt(method, info, &mut scoped);
                    }
                    ClassMember::Constructor { params, body, .. } => {
                        let mut scoped = std::collections::HashMap::new();
                        for param in params {
                            if let Some(type_hint) = &param.type_hint {
                                scoped.insert(param.name.to_lowercase(), type_hint.to_lowercase());
                            }
                        }
                        lower_pascal_method_pointer_body(body, info, &mut scoped);
                    }
                    _ => {}
                }
            }
        }
        StmtKind::Block(body) => {
            let mut scoped = var_types.clone();
            lower_pascal_method_pointer_body(body, info, &mut scoped);
        }
        StmtKind::If { cond, then_body, elifs, else_body } => {
            lower_pascal_method_pointer_expr(cond, info, var_types);
            for stmt in then_body {
                lower_pascal_method_pointer_stmt(stmt, info, &mut var_types.clone());
            }
            for (cond, body) in elifs {
                lower_pascal_method_pointer_expr(cond, info, var_types);
                for stmt in body {
                    lower_pascal_method_pointer_stmt(stmt, info, &mut var_types.clone());
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    lower_pascal_method_pointer_stmt(stmt, info, &mut var_types.clone());
                }
            }
        }
        StmtKind::For { init, cond, update, body } => {
            let mut scoped = var_types.clone();
            if let Some(init) = init {
                lower_pascal_method_pointer_stmt(init, info, &mut scoped);
            }
            if let Some(cond) = cond {
                lower_pascal_method_pointer_expr(cond, info, &scoped);
            }
            if let Some(update) = update {
                lower_pascal_method_pointer_expr(update, info, &scoped);
            }
            lower_pascal_method_pointer_body(body, info, &mut scoped);
        }
        StmtKind::While { cond, body, .. } | StmtKind::DoWhile { cond, body, .. } => {
            lower_pascal_method_pointer_expr(cond, info, var_types);
            lower_pascal_method_pointer_body(body, info, &mut var_types.clone());
        }
        _ => {}
    }
}

fn lower_pascal_method_pointer_expr(
    expr: &mut Expression,
    info: &PascalMethodPointerInfo,
    var_types: &std::collections::HashMap<String, String>,
) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            lower_pascal_method_pointer_expr(callee, info, var_types);
            let params = match &callee.kind {
                ExprKind::Ident(name) => info.routines.get(&name.to_lowercase()).cloned(),
                _ => None,
            };
            for (idx, arg) in args.iter_mut().enumerate() {
                lower_pascal_method_pointer_expr(&mut arg.value, info, var_types);
                if params
                    .as_ref()
                    .and_then(|params| params.get(idx))
                    .and_then(|param| param.type_hint.as_deref())
                    .is_some_and(|ty| is_pascal_procedural_type(ty, info))
                {
                    if let Some(lambda) = pascal_method_pointer_lambda(&arg.value, info, var_types) {
                        arg.value = lambda;
                    }
                }
            }
        }
        ExprKind::Member { object, .. } => lower_pascal_method_pointer_expr(object, info, var_types),
        ExprKind::Binary { left, right, .. } => {
            lower_pascal_method_pointer_expr(left, info, var_types);
            lower_pascal_method_pointer_expr(right, info, var_types);
        }
        ExprKind::Unary { expr, .. } => lower_pascal_method_pointer_expr(expr, info, var_types),
        ExprKind::Ternary { cond, then, else_ } => {
            lower_pascal_method_pointer_expr(cond, info, var_types);
            lower_pascal_method_pointer_expr(then, info, var_types);
            lower_pascal_method_pointer_expr(else_, info, var_types);
        }
        ExprKind::Index { object, index, .. } => {
            lower_pascal_method_pointer_expr(object, info, var_types);
            lower_pascal_method_pointer_expr(index, info, var_types);
        }
        ExprKind::Array(items) => {
            for item in items {
                if let Some(key) = &mut item.key {
                    lower_pascal_method_pointer_expr(key, info, var_types);
                }
                lower_pascal_method_pointer_expr(&mut item.value, info, var_types);
            }
        }
        _ => {}
    }
}

fn pascal_method_pointer_target_type(
    target: &Expression,
    info: &PascalMethodPointerInfo,
    var_types: &std::collections::HashMap<String, String>,
) -> Option<String> {
    match &target.kind {
        ExprKind::Ident(name) => var_types.get(&name.to_lowercase()).cloned(),
        ExprKind::Member { object, field, .. } => {
            let object_type = pascal_method_pointer_expr_type(object, info, var_types)?;
            info.fields
                .get(&object_type)?
                .get(&field.to_lowercase())
                .cloned()
        }
        _ => None,
    }
}

fn pascal_method_pointer_expr_type(
    expr: &Expression,
    info: &PascalMethodPointerInfo,
    var_types: &std::collections::HashMap<String, String>,
) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => var_types
            .get(&name.to_lowercase())
            .map(|ty| bare_type_name(ty).to_lowercase()),
        ExprKind::Member { object, field, .. } => {
            let object_type = pascal_method_pointer_expr_type(object, info, var_types)?;
            info.fields
                .get(&object_type)?
                .get(&field.to_lowercase())
                .map(|ty| bare_type_name(ty).to_lowercase())
        }
        _ => None,
    }
}

fn is_pascal_procedural_type(type_hint: &str, info: &PascalMethodPointerInfo) -> bool {
    let lower = type_hint.to_lowercase();
    lower.starts_with("procedure")
        || lower.starts_with("function")
        || info
            .aliases
            .get(&lower)
            .is_some_and(|aliased| is_pascal_procedural_type(aliased, info))
}

fn pascal_method_pointer_lambda(
    expr: &Expression,
    info: &PascalMethodPointerInfo,
    var_types: &std::collections::HashMap<String, String>,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &expr.kind else {
        return None;
    };
    let receiver_type = pascal_method_pointer_expr_type(object, info, var_types)?;
    if !pascal_type_has_zero_arg_method(&receiver_type, field, info) {
        return None;
    }
    Some(Expression::new(ExprKind::Lambda {
        params: Vec::new(),
        body: LambdaBody::Block(vec![Statement::new(StmtKind::Expr(Expression::new(
            ExprKind::Call {
                callee: Box::new(expr.clone()),
                args: Vec::new(),
                optional: false,
            },
        )))]),
        is_async: false,
        captures: Vec::new(),
    }))
}

fn pascal_type_has_zero_arg_method(
    type_name: &str,
    method_name: &str,
    info: &PascalMethodPointerInfo,
) -> bool {
    let lower = bare_type_name(type_name).to_lowercase();
    if info
        .methods
        .get(&lower)
        .is_some_and(|methods| methods.contains(&method_name.to_lowercase()))
    {
        return true;
    }
    info.parents
        .get(&lower)
        .into_iter()
        .flatten()
        .any(|parent| pascal_type_has_zero_arg_method(parent, method_name, info))
}

fn pascal_array_element_type(type_hint: &str) -> Option<String> {
    let trimmed = type_hint.trim();
    let lower = trimmed.to_ascii_lowercase();
    if !lower.starts_with("array") {
        return None;
    }
    let marker = " of ";
    let idx = lower.rfind(marker)?;
    Some(trimmed[idx + marker.len()..].trim().to_string())
}

fn collect_record_array_types(
    body: &[Statement],
    struct_names: &std::collections::HashSet<String>,
) -> std::collections::HashMap<String, String> {
    let mut out = std::collections::HashMap::new();
    for stmt in body {
        collect_record_array_types_stmt(stmt, struct_names, &mut out);
    }
    out
}

fn collect_record_array_types_stmt(
    stmt: &Statement,
    struct_names: &std::collections::HashSet<String>,
    out: &mut std::collections::HashMap<String, String>,
) {
    match &stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                let Some(type_hint) = decl.type_hint.as_deref() else {
                    continue;
                };
                let Some(element_type) = pascal_array_element_type(type_hint) else {
                    continue;
                };
                if struct_names.contains(&bare_type_name(&element_type).to_lowercase()) {
                    if let BindingPattern::Ident(name) = &decl.pattern {
                        out.insert(
                            name.to_lowercase(),
                            bare_type_name(&element_type).to_string(),
                        );
                    }
                }
            }
        }
        StmtKind::FunctionDecl { body, .. } | StmtKind::Block(body) => {
            for stmt in body {
                collect_record_array_types_stmt(stmt, struct_names, out);
            }
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                collect_record_array_types_member(member, struct_names, out);
            }
        }
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            for stmt in then_body {
                collect_record_array_types_stmt(stmt, struct_names, out);
            }
            for (_, body) in elifs {
                for stmt in body {
                    collect_record_array_types_stmt(stmt, struct_names, out);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    collect_record_array_types_stmt(stmt, struct_names, out);
                }
            }
        }
        StmtKind::For { body, .. }
        | StmtKind::ForIn { body, .. }
        | StmtKind::While { body, .. }
        | StmtKind::DoWhile { body, .. } => {
            for stmt in body {
                collect_record_array_types_stmt(stmt, struct_names, out);
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            for stmt in body {
                collect_record_array_types_stmt(stmt, struct_names, out);
            }
            for catch in catches {
                for stmt in &catch.body {
                    collect_record_array_types_stmt(stmt, struct_names, out);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    collect_record_array_types_stmt(stmt, struct_names, out);
                }
            }
            if let Some(body) = finally {
                for stmt in body {
                    collect_record_array_types_stmt(stmt, struct_names, out);
                }
            }
        }
        _ => {}
    }
}

fn collect_record_array_types_member(
    member: &ClassMember,
    struct_names: &std::collections::HashSet<String>,
    out: &mut std::collections::HashMap<String, String>,
) {
    match member {
        ClassMember::Field {
            name, type_hint, ..
        } => {
            let Some(type_hint) = type_hint.as_deref() else {
                return;
            };
            let Some(element_type) = pascal_array_element_type(type_hint) else {
                return;
            };
            if struct_names.contains(&bare_type_name(&element_type).to_lowercase()) {
                out.insert(
                    name.to_lowercase(),
                    bare_type_name(&element_type).to_string(),
                );
            }
        }
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            collect_record_array_types_stmt(stmt, struct_names, out);
        }
        ClassMember::Constructor { body, .. } => {
            for stmt in body {
                collect_record_array_types_stmt(stmt, struct_names, out);
            }
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                for stmt in getter {
                    collect_record_array_types_stmt(stmt, struct_names, out);
                }
            }
            if let Some(setter) = setter {
                for stmt in &setter.body {
                    collect_record_array_types_stmt(stmt, struct_names, out);
                }
            }
        }
        _ => {}
    }
}

fn rewrite_bare_parameterless_method_refs(body: &mut [Statement]) {
    for stmt in body {
        match &mut stmt.kind {
            StmtKind::ClassDecl { members, .. } | StmtKind::StructDecl { members, .. } => {
                let method_names: std::collections::HashSet<String> = members
                    .iter()
                    .filter_map(|member| {
                        let ClassMember::Method(stmt) = member else {
                            return None;
                        };
                        let StmtKind::FunctionDecl { name, params, .. } = &stmt.kind else {
                            return None;
                        };
                        params.is_empty().then(|| name.to_lowercase())
                    })
                    .collect();
                if method_names.is_empty() {
                    continue;
                }
                for member in members {
                    rewrite_bare_parameterless_method_refs_member(member, &method_names);
                }
            }
            StmtKind::ModuleDecl { members, .. } => {
                for member in members {
                    if let ClassMember::NestedType(stmt) | ClassMember::Method(stmt) = member {
                        rewrite_bare_parameterless_method_refs(std::slice::from_mut(stmt));
                    }
                }
            }
            _ => {}
        }
    }
}

fn rewrite_bare_parameterless_method_refs_member(
    member: &mut ClassMember,
    method_names: &std::collections::HashSet<String>,
) {
    match member {
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            rewrite_bare_parameterless_method_refs_stmt(stmt, method_names);
        }
        ClassMember::Constructor { params, body, .. } => {
            let scoped = method_names_without_params(method_names, params);
            for stmt in body {
                rewrite_bare_parameterless_method_refs_stmt(stmt, &scoped);
            }
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                for stmt in getter {
                    rewrite_bare_parameterless_method_refs_stmt(stmt, method_names);
                }
            }
            if let Some(setter) = setter {
                for stmt in &mut setter.body {
                    rewrite_bare_parameterless_method_refs_stmt(stmt, method_names);
                }
            }
        }
        _ => {}
    }
}

fn rewrite_bare_parameterless_method_refs_stmt(
    stmt: &mut Statement,
    method_names: &std::collections::HashSet<String>,
) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) => rewrite_bare_parameterless_method_refs_expr(expr, method_names),
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_bare_parameterless_method_refs_expr(init, method_names);
                }
            }
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_bare_parameterless_method_refs_expr(target, method_names);
            }
            rewrite_bare_parameterless_method_refs_expr(value, method_names);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_bare_parameterless_method_refs_expr(target, method_names);
            rewrite_bare_parameterless_method_refs_expr(value, method_names);
        }
        StmtKind::FunctionDecl { params, body, .. } => {
            let scoped = method_names_without_params(method_names, params);
            for stmt in body {
                rewrite_bare_parameterless_method_refs_stmt(stmt, &scoped);
            }
        }
        StmtKind::Block(body) => {
            for stmt in body {
                rewrite_bare_parameterless_method_refs_stmt(stmt, method_names);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_bare_parameterless_method_refs_expr(cond, method_names);
            for stmt in then_body {
                rewrite_bare_parameterless_method_refs_stmt(stmt, method_names);
            }
            for (cond, body) in elifs {
                rewrite_bare_parameterless_method_refs_expr(cond, method_names);
                for stmt in body {
                    rewrite_bare_parameterless_method_refs_stmt(stmt, method_names);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_bare_parameterless_method_refs_stmt(stmt, method_names);
                }
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                rewrite_bare_parameterless_method_refs_stmt(init, method_names);
            }
            if let Some(cond) = cond {
                rewrite_bare_parameterless_method_refs_expr(cond, method_names);
            }
            if let Some(update) = update {
                rewrite_bare_parameterless_method_refs_expr(update, method_names);
            }
            for stmt in body {
                rewrite_bare_parameterless_method_refs_stmt(stmt, method_names);
            }
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            rewrite_bare_parameterless_method_refs_expr(iter, method_names);
            for stmt in body {
                rewrite_bare_parameterless_method_refs_stmt(stmt, method_names);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_bare_parameterless_method_refs_stmt(stmt, method_names);
                }
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            rewrite_bare_parameterless_method_refs_expr(cond, method_names);
            for stmt in body {
                rewrite_bare_parameterless_method_refs_stmt(stmt, method_names);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_bare_parameterless_method_refs_stmt(stmt, method_names);
                }
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            for stmt in body {
                rewrite_bare_parameterless_method_refs_stmt(stmt, method_names);
            }
            rewrite_bare_parameterless_method_refs_expr(cond, method_names);
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            for stmt in body {
                rewrite_bare_parameterless_method_refs_stmt(stmt, method_names);
            }
            for catch in catches {
                for stmt in &mut catch.body {
                    rewrite_bare_parameterless_method_refs_stmt(stmt, method_names);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_bare_parameterless_method_refs_stmt(stmt, method_names);
                }
            }
            if let Some(body) = finally {
                for stmt in body {
                    rewrite_bare_parameterless_method_refs_stmt(stmt, method_names);
                }
            }
        }
        StmtKind::With { items, body, .. } => {
            for item in items {
                rewrite_bare_parameterless_method_refs_expr(&mut item.expr, method_names);
            }
            for stmt in body {
                rewrite_bare_parameterless_method_refs_stmt(stmt, method_names);
            }
        }
        StmtKind::Return(Some(expr))
        | StmtKind::Throw {
            expr: Some(expr), ..
        } => {
            rewrite_bare_parameterless_method_refs_expr(expr, method_names);
        }
        _ => {}
    }
}

fn rewrite_bare_parameterless_method_refs_expr(
    expr: &mut Expression,
    method_names: &std::collections::HashSet<String>,
) {
    match &mut expr.kind {
        ExprKind::Ident(name) if method_names.contains(&name.to_lowercase()) => {
            let callee = Expression::new(ExprKind::Member {
                object: Box::new(Expression::new(ExprKind::This)),
                field: name.clone(),
                null_safe: false,
            });
            expr.kind = ExprKind::Call {
                callee: Box::new(callee),
                args: Vec::new(),
                optional: false,
            };
        }
        ExprKind::Call { args, .. } => {
            for arg in args {
                rewrite_bare_parameterless_method_refs_expr(&mut arg.value, method_names);
            }
        }
        ExprKind::Member { object, .. } => {
            rewrite_bare_parameterless_method_refs_expr(object, method_names);
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_bare_parameterless_method_refs_expr(object, method_names);
            rewrite_bare_parameterless_method_refs_expr(index, method_names);
        }
        ExprKind::Binary { left, right, .. } => {
            rewrite_bare_parameterless_method_refs_expr(left, method_names);
            rewrite_bare_parameterless_method_refs_expr(right, method_names);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Await(expr)
        | ExprKind::Yield(Some(expr))
        | ExprKind::Spread(expr)
        | ExprKind::TypeOf(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Cast { expr, .. } => {
            rewrite_bare_parameterless_method_refs_expr(expr, method_names);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_bare_parameterless_method_refs_expr(cond, method_names);
            rewrite_bare_parameterless_method_refs_expr(then, method_names);
            rewrite_bare_parameterless_method_refs_expr(else_, method_names);
        }
        ExprKind::New { class, args } => {
            rewrite_bare_parameterless_method_refs_expr(class, method_names);
            for arg in args {
                rewrite_bare_parameterless_method_refs_expr(&mut arg.value, method_names);
            }
        }
        ExprKind::Array(elements) => {
            for element in elements {
                if let Some(key) = &mut element.key {
                    rewrite_bare_parameterless_method_refs_expr(key, method_names);
                }
                rewrite_bare_parameterless_method_refs_expr(&mut element.value, method_names);
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) => {
            for item in items {
                rewrite_bare_parameterless_method_refs_expr(item, method_names);
            }
        }
        ExprKind::NullCoalesce { left, right } => {
            rewrite_bare_parameterless_method_refs_expr(left, method_names);
            rewrite_bare_parameterless_method_refs_expr(right, method_names);
        }
        ExprKind::Range { start, end, .. } => {
            rewrite_bare_parameterless_method_refs_expr(start, method_names);
            rewrite_bare_parameterless_method_refs_expr(end, method_names);
        }
        ExprKind::Assign { target, value } | ExprKind::Walrus { target, value } => {
            rewrite_bare_parameterless_method_refs_expr(target, method_names);
            rewrite_bare_parameterless_method_refs_expr(value, method_names);
        }
        ExprKind::SuperCall { args, .. } => {
            for arg in args {
                rewrite_bare_parameterless_method_refs_expr(&mut arg.value, method_names);
            }
        }
        ExprKind::StaticAccess { class, member } => {
            rewrite_bare_parameterless_method_refs_expr(class, method_names);
            rewrite_bare_parameterless_method_refs_expr(member, method_names);
        }
        _ => {}
    }
}

#[derive(Clone)]
struct PascalIndexedPropertyInfo {
    name: String,
    getter: Option<String>,
    setter: Option<String>,
    is_default: bool,
}

type PascalIndexedPropertyMap = std::collections::HashMap<String, Vec<PascalIndexedPropertyInfo>>;

#[derive(Clone)]
struct PascalStaticPropertyInfo {
    name: String,
    getter: Option<String>,
    setter: Option<String>,
}

type PascalStaticPropertyMap = std::collections::HashMap<String, Vec<PascalStaticPropertyInfo>>;

fn collect_pascal_indexed_properties(body: &[Statement]) -> PascalIndexedPropertyMap {
    let mut out = std::collections::HashMap::new();
    let mut parents: std::collections::HashMap<String, Vec<String>> =
        std::collections::HashMap::new();
    for stmt in body {
        if let StmtKind::InterfaceDecl { name, members, .. } = &stmt.kind {
            let props: Vec<PascalIndexedPropertyInfo> = members
                .iter()
                .filter_map(|member| {
                    let InterfaceMember::Property { name, .. } = member else {
                        return None;
                    };
                    Some(PascalIndexedPropertyInfo {
                        name: name.to_lowercase(),
                        getter: Some(format!("Get{}", name)),
                        setter: Some(format!("Set{}", name)),
                        is_default: false,
                    })
                })
                .collect();
            if !props.is_empty() {
                out.insert(name.to_lowercase(), props);
            }
            continue;
        }
        let (name, class_parents, members) = match &stmt.kind {
            StmtKind::ClassDecl {
                name,
                parents,
                members,
                ..
            } => (name, parents.as_slice(), members),
            StmtKind::StructDecl { name, members, .. } => (name, &[][..], members),
            _ => continue,
        };
        parents.insert(
            name.to_lowercase(),
            class_parents
                .iter()
                .map(|parent| parent.to_lowercase())
                .collect(),
        );
        let mut props = Vec::new();
        for member in members {
            let ClassMember::Property {
                name,
                getter,
                setter,
                modifiers,
                ..
            } = member
            else {
                continue;
            };
            let mut is_indexed = false;
            let mut is_default = false;
            for decorator in &modifiers.decorators {
                if let ExprKind::Lit(Literal::Str(marker)) = &decorator.kind {
                    is_indexed |= marker == "__pascal_indexed_property";
                    is_default |= marker == "__pascal_default_property";
                }
            }
            if !is_indexed {
                continue;
            }
            props.push(PascalIndexedPropertyInfo {
                name: name.to_lowercase(),
                getter: getter
                    .as_ref()
                    .and_then(|body| property_body_accessor_name(body)),
                setter: setter
                    .as_ref()
                    .and_then(|setter| property_body_setter_name(&setter.body)),
                is_default,
            });
        }
        if !props.is_empty() {
            out.insert(name.to_lowercase(), props);
        }
    }
    let keys: Vec<String> = parents.keys().cloned().collect();
    for class_name in keys {
        inherit_pascal_indexed_properties(&class_name, &parents, &mut out);
    }
    out
}

fn collect_pascal_static_properties(body: &[Statement]) -> PascalStaticPropertyMap {
    let mut out = std::collections::HashMap::new();
    for stmt in body {
        let (StmtKind::ClassDecl { name, members, .. }
        | StmtKind::StructDecl { name, members, .. }) = &stmt.kind
        else {
            continue;
        };
        let mut props = Vec::new();
        for member in members {
            let ClassMember::Property {
                name,
                getter,
                setter,
                modifiers,
                ..
            } = member
            else {
                continue;
            };
            if !modifiers.is_static {
                continue;
            }
            props.push(PascalStaticPropertyInfo {
                name: name.to_lowercase(),
                getter: getter
                    .as_ref()
                    .and_then(|body| property_body_accessor_name(body)),
                setter: setter
                    .as_ref()
                    .and_then(|setter| property_body_setter_name(&setter.body)),
            });
        }
        if !props.is_empty() {
            out.insert(name.to_lowercase(), props);
        }
    }
    out
}

fn rewrite_pascal_static_properties(body: &mut [Statement], properties: &PascalStaticPropertyMap) {
    for stmt in body {
        rewrite_pascal_static_properties_stmt(stmt, properties);
    }
}

fn rewrite_pascal_static_properties_stmt(
    stmt: &mut Statement,
    properties: &PascalStaticPropertyMap,
) {
    match &mut stmt.kind {
        StmtKind::Assign { targets, value } if targets.len() == 1 => {
            rewrite_pascal_static_properties_expr(value, properties);
            if let Some(call) =
                pascal_static_property_setter_call(&targets[0], value.clone(), properties)
            {
                stmt.kind = StmtKind::Expr(call);
            } else {
                rewrite_pascal_static_properties_expr(&mut targets[0], properties);
            }
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_pascal_static_properties_expr(target, properties);
            }
            rewrite_pascal_static_properties_expr(value, properties);
        }
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            rewrite_pascal_static_properties_expr(expr, properties);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_pascal_static_properties_expr(init, properties);
                }
            }
        }
        StmtKind::FunctionDecl { body, .. } | StmtKind::Block(body) => {
            for stmt in body {
                rewrite_pascal_static_properties_stmt(stmt, properties);
            }
        }
        StmtKind::ClassDecl { members, .. } | StmtKind::StructDecl { members, .. } => {
            for member in members {
                rewrite_pascal_static_properties_member(member, properties);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_pascal_static_properties_expr(cond, properties);
            for stmt in then_body {
                rewrite_pascal_static_properties_stmt(stmt, properties);
            }
            for (cond, body) in elifs {
                rewrite_pascal_static_properties_expr(cond, properties);
                for stmt in body {
                    rewrite_pascal_static_properties_stmt(stmt, properties);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_pascal_static_properties_stmt(stmt, properties);
                }
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                rewrite_pascal_static_properties_stmt(init, properties);
            }
            if let Some(cond) = cond {
                rewrite_pascal_static_properties_expr(cond, properties);
            }
            if let Some(update) = update {
                rewrite_pascal_static_properties_expr(update, properties);
            }
            for stmt in body {
                rewrite_pascal_static_properties_stmt(stmt, properties);
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            rewrite_pascal_static_properties_expr(cond, properties);
            for stmt in body {
                rewrite_pascal_static_properties_stmt(stmt, properties);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_pascal_static_properties_stmt(stmt, properties);
                }
            }
        }
        _ => {}
    }
}

fn rewrite_pascal_static_properties_member(
    member: &mut ClassMember,
    properties: &PascalStaticPropertyMap,
) {
    match member {
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            rewrite_pascal_static_properties_stmt(stmt, properties)
        }
        ClassMember::Constructor { body, .. } => {
            for stmt in body {
                rewrite_pascal_static_properties_stmt(stmt, properties);
            }
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                for stmt in getter {
                    rewrite_pascal_static_properties_stmt(stmt, properties);
                }
            }
            if let Some(setter) = setter {
                for stmt in &mut setter.body {
                    rewrite_pascal_static_properties_stmt(stmt, properties);
                }
            }
        }
        _ => {}
    }
}

fn rewrite_pascal_static_properties_expr(
    expr: &mut Expression,
    properties: &PascalStaticPropertyMap,
) {
    match &mut expr.kind {
        ExprKind::Member { object, field, .. } => {
            rewrite_pascal_static_properties_expr(object, properties);
            if let ExprKind::Ident(class_name) = &object.kind {
                if let Some(getter) = properties
                    .get(&class_name.to_lowercase())
                    .and_then(|props| {
                        props
                            .iter()
                            .find(|prop| prop.name.eq_ignore_ascii_case(field))
                    })
                    .and_then(|prop| prop.getter.clone())
                {
                    *expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Member {
                            object: Box::new(Expression::ident(class_name)),
                            field: getter,
                            null_safe: false,
                        })),
                        args: Vec::new(),
                        optional: false,
                    });
                }
            }
        }
        ExprKind::Call { callee, args, .. } => {
            rewrite_pascal_static_properties_expr(callee, properties);
            for arg in args {
                rewrite_pascal_static_properties_expr(&mut arg.value, properties);
            }
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_pascal_static_properties_expr(object, properties);
            rewrite_pascal_static_properties_expr(index, properties);
        }
        ExprKind::Binary { left, right, .. } => {
            rewrite_pascal_static_properties_expr(left, properties);
            rewrite_pascal_static_properties_expr(right, properties);
        }
        ExprKind::Unary { expr, .. } => rewrite_pascal_static_properties_expr(expr, properties),
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_pascal_static_properties_expr(cond, properties);
            rewrite_pascal_static_properties_expr(then, properties);
            rewrite_pascal_static_properties_expr(else_, properties);
        }
        _ => {}
    }
}

fn pascal_static_property_setter_call(
    target: &Expression,
    value: Expression,
    properties: &PascalStaticPropertyMap,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &target.kind else {
        return None;
    };
    let ExprKind::Ident(class_name) = &object.kind else {
        return None;
    };
    let setter = properties
        .get(&class_name.to_lowercase())?
        .iter()
        .find(|prop| prop.name.eq_ignore_ascii_case(field))?
        .setter
        .clone()?;
    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident(class_name)),
            field: setter,
            null_safe: false,
        })),
        args: vec![Argument::positional(value)],
        optional: false,
    }))
}

fn inherit_pascal_indexed_properties(
    class_name: &str,
    parents: &std::collections::HashMap<String, Vec<String>>,
    properties: &mut PascalIndexedPropertyMap,
) -> Vec<PascalIndexedPropertyInfo> {
    let mut inherited = Vec::new();
    for parent in parents.get(class_name).cloned().unwrap_or_default() {
        inherited.extend(inherit_pascal_indexed_properties(
            &parent, parents, properties,
        ));
        inherited.extend(properties.get(&parent).cloned().unwrap_or_default());
    }
    if !inherited.is_empty() {
        let props = properties.entry(class_name.to_string()).or_default();
        for prop in &inherited {
            if !props.iter().any(|existing| existing.name == prop.name) {
                props.push(prop.clone());
            }
        }
    }
    properties.get(class_name).cloned().unwrap_or_default()
}

fn property_body_accessor_name(body: &[Statement]) -> Option<String> {
    let [stmt] = body else {
        return None;
    };
    let StmtKind::Return(Some(expr)) = &stmt.kind else {
        return None;
    };
    property_call_accessor_name(expr)
}

fn property_body_setter_name(body: &[Statement]) -> Option<String> {
    let [stmt] = body else {
        return None;
    };
    let StmtKind::Expr(expr) = &stmt.kind else {
        return None;
    };
    property_call_accessor_name(expr)
}

fn property_call_accessor_name(expr: &Expression) -> Option<String> {
    let ExprKind::Call { callee, .. } = &expr.kind else {
        return None;
    };
    match &callee.kind {
        ExprKind::Member { field, .. } => Some(field.clone()),
        ExprKind::Ident(name) => Some(name.clone()),
        _ => None,
    }
}

fn rewrite_pascal_indexed_properties(
    body: &mut [Statement],
    properties: &PascalIndexedPropertyMap,
) {
    let mut var_types = std::collections::HashMap::new();
    for stmt in body.iter() {
        collect_pascal_var_types_stmt(stmt, &mut var_types);
    }
    for stmt in body {
        rewrite_pascal_indexed_properties_stmt(stmt, properties, &mut var_types, None);
    }
}

fn collect_pascal_var_types_stmt(
    stmt: &Statement,
    out: &mut std::collections::HashMap<String, String>,
) {
    if let StmtKind::VarDecl { declarations, .. } = &stmt.kind {
        for decl in declarations {
            if let (BindingPattern::Ident(name), Some(type_hint)) = (&decl.pattern, &decl.type_hint)
            {
                out.insert(
                    name.to_lowercase(),
                    bare_type_name(type_hint).to_lowercase(),
                );
            }
        }
    }
}

fn rewrite_pascal_indexed_properties_stmt(
    stmt: &mut Statement,
    properties: &PascalIndexedPropertyMap,
    var_types: &mut std::collections::HashMap<String, String>,
    current_class: Option<&str>,
) {
    match &mut stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations.iter_mut() {
                if let Some(init) = &mut decl.init {
                    rewrite_pascal_indexed_properties_expr(
                        init,
                        properties,
                        var_types,
                        current_class,
                    );
                }
            }
            for decl in declarations {
                if let (BindingPattern::Ident(name), Some(type_hint)) =
                    (&decl.pattern, &decl.type_hint)
                {
                    var_types.insert(
                        name.to_lowercase(),
                        bare_type_name(type_hint).to_lowercase(),
                    );
                }
            }
        }
        StmtKind::Assign { targets, value } if targets.len() == 1 => {
            rewrite_pascal_indexed_properties_expr(value, properties, var_types, current_class);
            if let Some(call) = pascal_indexed_property_setter_call(
                &targets[0],
                value.clone(),
                properties,
                var_types,
                current_class,
            ) {
                stmt.kind = StmtKind::Expr(call);
            } else {
                rewrite_pascal_indexed_properties_expr(
                    &mut targets[0],
                    properties,
                    var_types,
                    current_class,
                );
            }
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_pascal_indexed_properties_expr(
                    target,
                    properties,
                    var_types,
                    current_class,
                );
            }
            rewrite_pascal_indexed_properties_expr(value, properties, var_types, current_class);
        }
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            rewrite_pascal_indexed_properties_expr(expr, properties, var_types, current_class);
        }
        StmtKind::FunctionDecl { params, body, .. } => {
            let mut scoped = var_types.clone();
            for param in params {
                if let Some(type_hint) = &param.type_hint {
                    scoped.insert(
                        param.name.to_lowercase(),
                        bare_type_name(type_hint).to_lowercase(),
                    );
                }
            }
            for stmt in body {
                rewrite_pascal_indexed_properties_stmt(
                    stmt,
                    properties,
                    &mut scoped,
                    current_class,
                );
            }
        }
        StmtKind::ClassDecl { name, members, .. } | StmtKind::StructDecl { name, members, .. } => {
            for member in members {
                rewrite_pascal_indexed_properties_member(member, properties, name);
            }
        }
        StmtKind::Block(body) => {
            let mut scoped = var_types.clone();
            for stmt in body {
                rewrite_pascal_indexed_properties_stmt(
                    stmt,
                    properties,
                    &mut scoped,
                    current_class,
                );
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_pascal_indexed_properties_expr(cond, properties, var_types, current_class);
            for stmt in then_body {
                rewrite_pascal_indexed_properties_stmt(
                    stmt,
                    properties,
                    &mut var_types.clone(),
                    current_class,
                );
            }
            for (cond, body) in elifs {
                rewrite_pascal_indexed_properties_expr(cond, properties, var_types, current_class);
                for stmt in body {
                    rewrite_pascal_indexed_properties_stmt(
                        stmt,
                        properties,
                        &mut var_types.clone(),
                        current_class,
                    );
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_pascal_indexed_properties_stmt(
                        stmt,
                        properties,
                        &mut var_types.clone(),
                        current_class,
                    );
                }
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            let mut scoped = var_types.clone();
            if let Some(init) = init {
                rewrite_pascal_indexed_properties_stmt(
                    init,
                    properties,
                    &mut scoped,
                    current_class,
                );
            }
            if let Some(cond) = cond {
                rewrite_pascal_indexed_properties_expr(
                    cond,
                    properties,
                    &mut scoped,
                    current_class,
                );
            }
            if let Some(update) = update {
                rewrite_pascal_indexed_properties_expr(
                    update,
                    properties,
                    &mut scoped,
                    current_class,
                );
            }
            for stmt in body {
                rewrite_pascal_indexed_properties_stmt(
                    stmt,
                    properties,
                    &mut scoped,
                    current_class,
                );
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            rewrite_pascal_indexed_properties_expr(cond, properties, var_types, current_class);
            for stmt in body {
                rewrite_pascal_indexed_properties_stmt(
                    stmt,
                    properties,
                    &mut var_types.clone(),
                    current_class,
                );
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_pascal_indexed_properties_stmt(
                        stmt,
                        properties,
                        &mut var_types.clone(),
                        current_class,
                    );
                }
            }
        }
        _ => {}
    }
}

fn rewrite_pascal_indexed_properties_member(
    member: &mut ClassMember,
    properties: &PascalIndexedPropertyMap,
    class_name: &str,
) {
    match member {
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            rewrite_pascal_indexed_properties_stmt(
                stmt,
                properties,
                &mut std::collections::HashMap::new(),
                Some(&class_name.to_lowercase()),
            );
        }
        ClassMember::Constructor { params, body, .. } => {
            let mut scoped = std::collections::HashMap::new();
            for param in params {
                if let Some(type_hint) = &param.type_hint {
                    scoped.insert(
                        param.name.to_lowercase(),
                        bare_type_name(type_hint).to_lowercase(),
                    );
                }
            }
            for stmt in body {
                rewrite_pascal_indexed_properties_stmt(
                    stmt,
                    properties,
                    &mut scoped,
                    Some(&class_name.to_lowercase()),
                );
            }
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                for stmt in getter {
                    rewrite_pascal_indexed_properties_stmt(
                        stmt,
                        properties,
                        &mut std::collections::HashMap::new(),
                        Some(&class_name.to_lowercase()),
                    );
                }
            }
            if let Some(setter) = setter {
                for stmt in &mut setter.body {
                    rewrite_pascal_indexed_properties_stmt(
                        stmt,
                        properties,
                        &mut std::collections::HashMap::new(),
                        Some(&class_name.to_lowercase()),
                    );
                }
            }
        }
        _ => {}
    }
}

fn rewrite_pascal_indexed_properties_expr(
    expr: &mut Expression,
    properties: &PascalIndexedPropertyMap,
    var_types: &std::collections::HashMap<String, String>,
    current_class: Option<&str>,
) {
    match &mut expr.kind {
        ExprKind::Index { object, index, .. } => {
            rewrite_pascal_indexed_properties_expr(object, properties, var_types, current_class);
            rewrite_pascal_indexed_properties_expr(index, properties, var_types, current_class);
            if let Some(call) = pascal_indexed_property_getter_call(
                object,
                (**index).clone(),
                properties,
                var_types,
                current_class,
            ) {
                *expr = call;
            }
        }
        ExprKind::Call { callee, args, .. } => {
            rewrite_pascal_indexed_properties_expr(callee, properties, var_types, current_class);
            for arg in args {
                rewrite_pascal_indexed_properties_expr(
                    &mut arg.value,
                    properties,
                    var_types,
                    current_class,
                );
            }
        }
        ExprKind::Member { object, .. } => {
            rewrite_pascal_indexed_properties_expr(object, properties, var_types, current_class);
            if let Some(call) =
                pascal_property_getter_call(expr, properties, var_types, current_class)
            {
                *expr = call;
            }
        }
        ExprKind::Binary { left, right, .. } => {
            rewrite_pascal_indexed_properties_expr(left, properties, var_types, current_class);
            rewrite_pascal_indexed_properties_expr(right, properties, var_types, current_class);
        }
        ExprKind::Unary { expr, .. } => {
            rewrite_pascal_indexed_properties_expr(expr, properties, var_types, current_class);
        }
        ExprKind::Array(elements) => {
            for element in elements {
                if let Some(key) = &mut element.key {
                    rewrite_pascal_indexed_properties_expr(
                        key,
                        properties,
                        var_types,
                        current_class,
                    );
                }
                rewrite_pascal_indexed_properties_expr(
                    &mut element.value,
                    properties,
                    var_types,
                    current_class,
                );
            }
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_pascal_indexed_properties_expr(cond, properties, var_types, current_class);
            rewrite_pascal_indexed_properties_expr(then, properties, var_types, current_class);
            rewrite_pascal_indexed_properties_expr(else_, properties, var_types, current_class);
        }
        _ => {}
    }
}

fn pascal_property_getter_call(
    expr: &Expression,
    properties: &PascalIndexedPropertyMap,
    var_types: &std::collections::HashMap<String, String>,
    current_class: Option<&str>,
) -> Option<Expression> {
    let ExprKind::Member {
        object,
        field,
        null_safe,
    } = &expr.kind
    else {
        return None;
    };
    let class_name = pascal_expr_class_name(object, var_types, current_class)?;
    let prop = properties
        .get(&class_name)?
        .iter()
        .find(|prop| prop.name.eq_ignore_ascii_case(field))?;
    let getter = prop.getter.as_ref()?;
    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: object.clone(),
            field: getter.clone(),
            null_safe: *null_safe,
        })),
        args: Vec::new(),
        optional: false,
    }))
}

fn pascal_indexed_property_getter_call(
    object: &Expression,
    index: Expression,
    properties: &PascalIndexedPropertyMap,
    var_types: &std::collections::HashMap<String, String>,
    current_class: Option<&str>,
) -> Option<Expression> {
    let (receiver, prop) =
        pascal_indexed_property_target(object, properties, var_types, current_class, true)?;
    let getter = prop.getter.as_ref()?;
    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(receiver),
            field: getter.clone(),
            null_safe: false,
        })),
        args: vec![Argument::positional(index)],
        optional: false,
    }))
}

fn pascal_indexed_property_setter_call(
    target: &Expression,
    value: Expression,
    properties: &PascalIndexedPropertyMap,
    var_types: &std::collections::HashMap<String, String>,
    current_class: Option<&str>,
) -> Option<Expression> {
    let ExprKind::Index { object, index, .. } = &target.kind else {
        return None;
    };
    let (receiver, prop) =
        pascal_indexed_property_target(object, properties, var_types, current_class, true)?;
    let setter = prop.setter.as_ref()?;
    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(receiver),
            field: setter.clone(),
            null_safe: false,
        })),
        args: vec![
            Argument::positional((**index).clone()),
            Argument::positional(value),
        ],
        optional: false,
    }))
}

fn pascal_indexed_property_target(
    object: &Expression,
    properties: &PascalIndexedPropertyMap,
    var_types: &std::collections::HashMap<String, String>,
    current_class: Option<&str>,
    allow_default: bool,
) -> Option<(Expression, PascalIndexedPropertyInfo)> {
    match &object.kind {
        ExprKind::Member {
            object: receiver,
            field,
            ..
        } => {
            let class_name = pascal_expr_class_name(receiver, var_types, current_class)?;
            let prop = properties
                .get(&class_name)?
                .iter()
                .find(|prop| prop.name.eq_ignore_ascii_case(field))?
                .clone();
            Some(((**receiver).clone(), prop))
        }
        ExprKind::Ident(name) => {
            if let Some(class_name) = current_class {
                if let Some(prop) = properties
                    .get(class_name)
                    .and_then(|props| {
                        props
                            .iter()
                            .find(|prop| prop.name.eq_ignore_ascii_case(name))
                    })
                    .cloned()
                {
                    return Some((Expression::new(ExprKind::This), prop));
                }
            }
            if allow_default {
                let class_name = pascal_expr_class_name(object, var_types, current_class)?;
                let prop = properties
                    .get(&class_name)?
                    .iter()
                    .find(|prop| prop.is_default)?
                    .clone();
                return Some((object.clone(), prop));
            }
            None
        }
        _ if allow_default => {
            let class_name = pascal_expr_class_name(object, var_types, current_class)?;
            let prop = properties
                .get(&class_name)?
                .iter()
                .find(|prop| prop.is_default)?
                .clone();
            Some((object.clone(), prop))
        }
        _ => None,
    }
}

fn pascal_expr_class_name(
    expr: &Expression,
    var_types: &std::collections::HashMap<String, String>,
    current_class: Option<&str>,
) -> Option<String> {
    match &expr.kind {
        ExprKind::This => current_class.map(str::to_string),
        ExprKind::Ident(name) => var_types.get(&name.to_lowercase()).cloned(),
        _ => None,
    }
}

fn method_names_without_params(
    method_names: &std::collections::HashSet<String>,
    params: &[Param],
) -> std::collections::HashSet<String> {
    let mut scoped = method_names.clone();
    for param in params {
        scoped.remove(&param.name.to_lowercase());
    }
    scoped
}

fn target_array_name(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => Some(name.to_lowercase()),
        ExprKind::Member { field, .. } => Some(field.to_lowercase()),
        _ => None,
    }
}

fn setlength_record_slot_assign(
    target: Expression,
    len: Expression,
    element_type: &str,
) -> Statement {
    let index = Expression::new(ExprKind::Binary {
        op: BinOp::Sub,
        left: Box::new(len),
        right: Box::new(Expression::int(1)),
    });
    Statement::new(StmtKind::Assign {
        targets: vec![Expression::new(ExprKind::Index {
            object: Box::new(target),
            index: Box::new(index),
            null_safe: false,
        })],
        value: Expression::new(ExprKind::New {
            class: Box::new(Expression::ident(element_type)),
            args: Vec::new(),
        }),
    })
}

fn materialize_record_array_setlength_stmt(
    stmt: &mut Statement,
    record_arrays: &std::collections::HashMap<String, String>,
) {
    match &mut stmt.kind {
        StmtKind::Block(body) => {
            materialize_record_array_setlength_body(body, record_arrays);
        }
        StmtKind::FunctionDecl { body, .. } => {
            materialize_record_array_setlength_body(body, record_arrays);
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                materialize_record_array_setlength_member(member, record_arrays);
            }
        }
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            materialize_record_array_setlength_body(then_body, record_arrays);
            for (_, body) in elifs {
                materialize_record_array_setlength_body(body, record_arrays);
            }
            if let Some(body) = else_body {
                materialize_record_array_setlength_body(body, record_arrays);
            }
        }
        StmtKind::For { body, .. }
        | StmtKind::ForIn { body, .. }
        | StmtKind::While { body, .. }
        | StmtKind::DoWhile { body, .. } => {
            materialize_record_array_setlength_body(body, record_arrays);
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            materialize_record_array_setlength_body(body, record_arrays);
            for catch in catches {
                materialize_record_array_setlength_body(&mut catch.body, record_arrays);
            }
            if let Some(body) = else_body {
                materialize_record_array_setlength_body(body, record_arrays);
            }
            if let Some(body) = finally {
                materialize_record_array_setlength_body(body, record_arrays);
            }
        }
        _ => {}
    }
}

fn materialize_record_array_setlength_body(
    body: &mut Vec<Statement>,
    record_arrays: &std::collections::HashMap<String, String>,
) {
    let mut i = 0;
    while i < body.len() {
        materialize_record_array_setlength_stmt(&mut body[i], record_arrays);
        let extra = match &body[i].kind {
            StmtKind::Expr(expr) => match &expr.kind {
                ExprKind::Call { callee, args, .. }
                    if args.len() == 2
                        && matches!(&callee.kind, ExprKind::Ident(name) if name.eq_ignore_ascii_case("SetLength")) =>
                {
                    target_array_name(&args[0].value)
                        .and_then(|name| record_arrays.get(&name))
                        .map(|element_type| {
                            setlength_record_slot_assign(
                                args[0].value.clone(),
                                args[1].value.clone(),
                                element_type,
                            )
                        })
                }
                _ => None,
            },
            _ => None,
        };
        if let Some(stmt) = extra {
            body.insert(i + 1, stmt);
            i += 2;
        } else {
            i += 1;
        }
    }
}

fn materialize_record_array_setlength_member(
    member: &mut ClassMember,
    record_arrays: &std::collections::HashMap<String, String>,
) {
    match member {
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            materialize_record_array_setlength_stmt(stmt, record_arrays);
        }
        ClassMember::Constructor { body, .. } => {
            materialize_record_array_setlength_body(body, record_arrays);
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                materialize_record_array_setlength_body(getter, record_arrays);
            }
            if let Some(setter) = setter {
                materialize_record_array_setlength_body(&mut setter.body, record_arrays);
            }
        }
        _ => {}
    }
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

fn collect_pascal_polymorphic_class_names(body: &[Statement]) -> std::collections::HashSet<String> {
    let mut names = std::collections::HashSet::new();
    for stmt in body {
        if let StmtKind::ClassDecl { parents, .. } = &stmt.kind {
            for parent in parents {
                names.insert(parent.to_lowercase());
            }
        }
    }
    names
}

fn erase_pascal_class_value_type_hints_stmt(
    stmt: &mut Statement,
    class_names: &std::collections::HashSet<String>,
) {
    match &mut stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if decl
                    .type_hint
                    .as_deref()
                    .is_some_and(|hint| class_names.contains(&bare_type_name(hint).to_lowercase()))
                {
                    decl.type_hint = None;
                }
                if let Some(init) = &mut decl.init {
                    erase_pascal_class_value_type_hints_expr(init, class_names);
                }
            }
        }
        StmtKind::FunctionDecl { params, body, .. } => {
            erase_pascal_class_param_type_hints(params, class_names);
            for stmt in body {
                erase_pascal_class_value_type_hints_stmt(stmt, class_names);
            }
        }
        StmtKind::Block(stmts) => {
            for stmt in stmts {
                erase_pascal_class_value_type_hints_stmt(stmt, class_names);
            }
        }
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            erase_pascal_class_value_type_hints_expr(expr, class_names)
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                erase_pascal_class_value_type_hints_expr(target, class_names);
            }
            erase_pascal_class_value_type_hints_expr(value, class_names);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            erase_pascal_class_value_type_hints_expr(target, class_names);
            erase_pascal_class_value_type_hints_expr(value, class_names);
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            erase_pascal_class_value_type_hints_expr(cond, class_names);
            for stmt in then_body {
                erase_pascal_class_value_type_hints_stmt(stmt, class_names);
            }
            for (cond, body) in elifs {
                erase_pascal_class_value_type_hints_expr(cond, class_names);
                for stmt in body {
                    erase_pascal_class_value_type_hints_stmt(stmt, class_names);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    erase_pascal_class_value_type_hints_stmt(stmt, class_names);
                }
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
            ..
        } => {
            erase_pascal_class_value_type_hints_expr(cond, class_names);
            for stmt in body {
                erase_pascal_class_value_type_hints_stmt(stmt, class_names);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    erase_pascal_class_value_type_hints_stmt(stmt, class_names);
                }
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
                erase_pascal_class_value_type_hints_stmt(init, class_names);
            }
            if let Some(cond) = cond {
                erase_pascal_class_value_type_hints_expr(cond, class_names);
            }
            if let Some(update) = update {
                erase_pascal_class_value_type_hints_expr(update, class_names);
            }
            for stmt in body {
                erase_pascal_class_value_type_hints_stmt(stmt, class_names);
            }
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            erase_pascal_class_value_type_hints_expr(iter, class_names);
            for stmt in body {
                erase_pascal_class_value_type_hints_stmt(stmt, class_names);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    erase_pascal_class_value_type_hints_stmt(stmt, class_names);
                }
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            for stmt in body {
                erase_pascal_class_value_type_hints_stmt(stmt, class_names);
            }
            erase_pascal_class_value_type_hints_expr(cond, class_names);
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            erase_pascal_class_value_type_hints_expr(expr, class_names);
            for case in cases {
                for cond in &mut case.conditions {
                    match cond {
                        CaseCondition::Value(expr) | CaseCondition::Comparison { expr, .. } => {
                            erase_pascal_class_value_type_hints_expr(expr, class_names);
                        }
                        CaseCondition::Range { from, to } => {
                            erase_pascal_class_value_type_hints_expr(from, class_names);
                            erase_pascal_class_value_type_hints_expr(to, class_names);
                        }
                    }
                }
                for stmt in &mut case.body {
                    erase_pascal_class_value_type_hints_stmt(stmt, class_names);
                }
            }
            if let Some(default) = default {
                for stmt in default {
                    erase_pascal_class_value_type_hints_stmt(stmt, class_names);
                }
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            for stmt in body {
                erase_pascal_class_value_type_hints_stmt(stmt, class_names);
            }
            for catch in catches {
                if let Some(when_clause) = &mut catch.when_clause {
                    erase_pascal_class_value_type_hints_expr(when_clause, class_names);
                }
                for stmt in &mut catch.body {
                    erase_pascal_class_value_type_hints_stmt(stmt, class_names);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    erase_pascal_class_value_type_hints_stmt(stmt, class_names);
                }
            }
            if let Some(body) = finally {
                for stmt in body {
                    erase_pascal_class_value_type_hints_stmt(stmt, class_names);
                }
            }
        }
        StmtKind::Using { resource, body, .. } => {
            erase_pascal_class_value_type_hints_expr(resource, class_names);
            for stmt in body {
                erase_pascal_class_value_type_hints_stmt(stmt, class_names);
            }
        }
        StmtKind::With { items, body, .. } => {
            for item in items {
                erase_pascal_class_value_type_hints_expr(&mut item.expr, class_names);
            }
            for stmt in body {
                erase_pascal_class_value_type_hints_stmt(stmt, class_names);
            }
        }
        StmtKind::ClassDecl { members, .. } | StmtKind::StructDecl { members, .. } => {
            for member in members {
                erase_pascal_class_value_type_hints_member(member, class_names);
            }
        }
        _ => {}
    }
}

fn erase_pascal_class_value_type_hints_member(
    member: &mut ClassMember,
    class_names: &std::collections::HashSet<String>,
) {
    match member {
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            erase_pascal_class_value_type_hints_stmt(stmt, class_names)
        }
        ClassMember::Constructor { params, body, .. } => {
            erase_pascal_class_param_type_hints(params, class_names);
            for stmt in body {
                erase_pascal_class_value_type_hints_stmt(stmt, class_names);
            }
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                for stmt in getter {
                    erase_pascal_class_value_type_hints_stmt(stmt, class_names);
                }
            }
            if let Some(setter) = setter {
                erase_pascal_class_param_type_hints(
                    std::slice::from_mut(&mut setter.param),
                    class_names,
                );
                for stmt in &mut setter.body {
                    erase_pascal_class_value_type_hints_stmt(stmt, class_names);
                }
            }
        }
        _ => {}
    }
}

fn erase_pascal_class_param_type_hints(
    params: &mut [Param],
    class_names: &std::collections::HashSet<String>,
) {
    for param in params {
        if param
            .type_hint
            .as_deref()
            .is_some_and(|hint| class_names.contains(&bare_type_name(hint).to_lowercase()))
        {
            param.type_hint = None;
        }
        if let Some(default) = &mut param.default {
            erase_pascal_class_value_type_hints_expr(default, class_names);
        }
    }
}

fn erase_pascal_class_value_type_hints_expr(
    expr: &mut Expression,
    class_names: &std::collections::HashSet<String>,
) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            erase_pascal_class_value_type_hints_expr(callee, class_names);
            for arg in args {
                erase_pascal_class_value_type_hints_expr(&mut arg.value, class_names);
            }
        }
        ExprKind::Member { object, .. } => {
            erase_pascal_class_value_type_hints_expr(object, class_names)
        }
        ExprKind::Index { object, index, .. } => {
            erase_pascal_class_value_type_hints_expr(object, class_names);
            erase_pascal_class_value_type_hints_expr(index, class_names);
        }
        ExprKind::Binary { left, right, .. } => {
            erase_pascal_class_value_type_hints_expr(left, class_names);
            erase_pascal_class_value_type_hints_expr(right, class_names);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Await(expr)
        | ExprKind::Yield(Some(expr))
        | ExprKind::YieldFrom(expr)
        | ExprKind::Spread(expr)
        | ExprKind::TypeOf(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::Cast { expr, .. } => {
            erase_pascal_class_value_type_hints_expr(expr, class_names)
        }
        ExprKind::Ternary { cond, then, else_ } => {
            erase_pascal_class_value_type_hints_expr(cond, class_names);
            erase_pascal_class_value_type_hints_expr(then, class_names);
            erase_pascal_class_value_type_hints_expr(else_, class_names);
        }
        ExprKind::New { class, args } => {
            erase_pascal_class_value_type_hints_expr(class, class_names);
            for arg in args {
                erase_pascal_class_value_type_hints_expr(&mut arg.value, class_names);
            }
        }
        ExprKind::Array(elements) => {
            for element in elements {
                if let Some(key) = &mut element.key {
                    erase_pascal_class_value_type_hints_expr(key, class_names);
                }
                erase_pascal_class_value_type_hints_expr(&mut element.value, class_names);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        erase_pascal_class_value_type_hints_expr(key, class_names);
                        erase_pascal_class_value_type_hints_expr(value, class_names);
                    }
                    ObjectProperty::Spread(value) => {
                        erase_pascal_class_value_type_hints_expr(value, class_names);
                    }
                    _ => {}
                }
            }
        }
        ExprKind::Tuple(items) => {
            for item in items {
                erase_pascal_class_value_type_hints_expr(item, class_names);
            }
        }
        ExprKind::Lambda { body, params, .. } => {
            erase_pascal_class_param_type_hints(params, class_names);
            match body {
                LambdaBody::Expr(expr) => {
                    erase_pascal_class_value_type_hints_expr(expr, class_names)
                }
                LambdaBody::Block(body) => {
                    for stmt in body {
                        erase_pascal_class_value_type_hints_stmt(stmt, class_names);
                    }
                }
            }
        }
        ExprKind::Assign { target, value } => {
            erase_pascal_class_value_type_hints_expr(target, class_names);
            erase_pascal_class_value_type_hints_expr(value, class_names);
        }
        _ => {}
    }
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

fn null_array_initializer(count: usize) -> Expression {
    let elements = (0..count)
        .map(|_| ArrayElement {
            key: None,
            value: Expression::null(),
            spread: false,
            by_ref: false,
        })
        .collect();
    Expression::new(ExprKind::Array(elements))
}

fn fixed_array_length(type_hint: &str) -> Option<usize> {
    let trimmed = type_hint.trim();
    let lower = trimmed.to_ascii_lowercase();
    let rest = lower.strip_prefix("array[")?;
    let close = rest.find(']')?;
    let bounds = &trimmed["array[".len().."array[".len() + close];
    let mut total = 1usize;
    for dim in bounds.split(',') {
        let (lo, hi) = dim.split_once("..")?;
        let lo = lo.trim().parse::<isize>().ok()?;
        let hi = hi.trim().parse::<isize>().ok()?;
        if hi < lo {
            return None;
        }
        total = total.checked_mul((hi - lo + 1) as usize)?;
    }
    Some(total)
}

fn default_array_init_for_type(type_hint: &str) -> Option<Expression> {
    let trimmed = type_hint.trim();
    let lower = trimmed.to_ascii_lowercase();
    if !lower.starts_with("array") {
        return None;
    }
    if let Some(count) = fixed_array_length(trimmed) {
        return Some(null_array_initializer(count));
    }
    if lower.starts_with("array[") {
        return None;
    }
    Some(Expression::new(ExprKind::Array(Vec::new())))
}

fn default_field_init_for_type(type_hint: &str) -> Option<Expression> {
    if is_pascal_set_type_hint(type_hint) {
        return Some(empty_pascal_set_expr());
    }
    if let Some(init) = default_array_init_for_type(type_hint) {
        return Some(init);
    }
    match bare_type_name(type_hint).to_ascii_lowercase().as_str() {
        "integer" | "int" | "longint" | "word" | "byte" | "smallint" | "shortint" | "cardinal"
        | "real" | "double" | "single" | "extended" | "currency" => Some(Expression::int(0)),
        "boolean" | "bool" => Some(Expression::bool(false)),
        "string" | "ansistring" | "unicodestring" | "widestring" | "char" => {
            Some(Expression::string(""))
        }
        _ => None,
    }
}

fn default_init_const_bounded_arrays(body: &mut [Statement]) {
    let mut consts = std::collections::HashMap::new();
    collect_pascal_int_consts(body, &mut consts);
    if !consts.is_empty() {
        default_init_const_bounded_arrays_body(body, &consts);
    }
}

fn collect_pascal_int_consts(
    body: &[Statement],
    out: &mut std::collections::HashMap<String, i64>,
) {
    for stmt in body {
        match &stmt.kind {
            StmtKind::VarDecl {
                declarations,
                kind: VarDeclKind::Const,
            } => {
                for decl in declarations {
                    if let (BindingPattern::Ident(name), Some(init)) = (&decl.pattern, &decl.init)
                    {
                        if let Some(value) = const_int_expr(init) {
                            out.insert(name.to_lowercase(), value);
                        }
                    }
                }
            }
            StmtKind::FunctionDecl { body, .. } | StmtKind::Block(body) => {
                collect_pascal_int_consts(body, out);
            }
            _ => {}
        }
    }
}

fn default_init_const_bounded_arrays_body(
    body: &mut [Statement],
    consts: &std::collections::HashMap<String, i64>,
) {
    for stmt in body {
        match &mut stmt.kind {
            StmtKind::VarDecl { declarations, .. } => {
                for decl in declarations {
                    if decl.init.is_some() {
                        continue;
                    }
                    let Some(type_hint) = decl.type_hint.as_deref() else {
                        continue;
                    };
                    if let Some(count) = fixed_array_length_with_consts(type_hint, consts) {
                        decl.init = Some(null_array_initializer(count));
                    }
                }
            }
            StmtKind::FunctionDecl { body, .. } | StmtKind::Block(body) => {
                default_init_const_bounded_arrays_body(body, consts);
            }
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                default_init_const_bounded_arrays_body(then_body, consts);
                for (_, body) in elifs {
                    default_init_const_bounded_arrays_body(body, consts);
                }
                if let Some(body) = else_body {
                    default_init_const_bounded_arrays_body(body, consts);
                }
            }
            StmtKind::For { body, .. }
            | StmtKind::ForIn { body, .. }
            | StmtKind::While { body, .. }
            | StmtKind::DoWhile { body, .. } => {
                default_init_const_bounded_arrays_body(body, consts);
            }
            _ => {}
        }
    }
}

fn fixed_array_length_with_consts(
    type_hint: &str,
    consts: &std::collections::HashMap<String, i64>,
) -> Option<usize> {
    let trimmed = type_hint.trim();
    let lower = trimmed.to_ascii_lowercase();
    let rest = lower.strip_prefix("array[")?;
    let close = rest.find(']')?;
    let bounds = &trimmed["array[".len().."array[".len() + close];
    let mut total = 1usize;
    for dim in bounds.split(',') {
        let (lo, hi) = dim.split_once("..")?;
        let lo = pascal_bound_int(lo.trim(), consts)?;
        let hi = pascal_bound_int(hi.trim(), consts)?;
        if hi < lo {
            return None;
        }
        total = total.checked_mul((hi - lo + 1) as usize)?;
    }
    Some(total)
}

fn pascal_bound_int(value: &str, consts: &std::collections::HashMap<String, i64>) -> Option<i64> {
    value
        .parse::<i64>()
        .ok()
        .or_else(|| consts.get(&value.to_lowercase()).copied())
}

fn collect_struct_fields(body: &[Statement]) -> std::collections::HashMap<String, Vec<String>> {
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

fn rewrite_zero_based_string_indexes_stmt(
    stmt: &mut Statement,
    string_vars: &mut std::collections::HashSet<String>,
    zero_based_loop_vars: &mut std::collections::HashSet<String>,
) {
    match &mut stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_zero_based_string_indexes_expr(init, string_vars, zero_based_loop_vars);
                }
                if decl
                    .type_hint
                    .as_deref()
                    .is_some_and(|hint| hint.eq_ignore_ascii_case("String"))
                {
                    if let BindingPattern::Ident(name) = &decl.pattern {
                        string_vars.insert(name.to_lowercase());
                    }
                }
            }
        }
        StmtKind::FunctionDecl { params, body, .. } => {
            let mut scoped_strings = std::collections::HashSet::new();
            for param in params {
                if param
                    .type_hint
                    .as_deref()
                    .is_some_and(|hint| hint.eq_ignore_ascii_case("String"))
                {
                    scoped_strings.insert(param.name.to_lowercase());
                }
            }
            let mut scoped_zero = std::collections::HashSet::new();
            for stmt in body {
                rewrite_zero_based_string_indexes_stmt(stmt, &mut scoped_strings, &mut scoped_zero);
            }
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                rewrite_zero_based_string_indexes_member(member);
            }
        }
        StmtKind::Block(body) | StmtKind::NamespaceDecl { body, .. } => {
            for stmt in body {
                rewrite_zero_based_string_indexes_stmt(stmt, string_vars, zero_based_loop_vars);
            }
        }
        StmtKind::Expr(expr) => {
            rewrite_zero_based_string_indexes_expr(expr, string_vars, zero_based_loop_vars);
        }
        StmtKind::Return(Some(expr))
        | StmtKind::Throw {
            expr: Some(expr), ..
        } => {
            rewrite_zero_based_string_indexes_expr(expr, string_vars, zero_based_loop_vars);
        }
        StmtKind::Throw {
            cause: Some(cause), ..
        } => rewrite_zero_based_string_indexes_expr(cause, string_vars, zero_based_loop_vars),
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_zero_based_string_indexes_expr(target, string_vars, zero_based_loop_vars);
            }
            rewrite_zero_based_string_indexes_expr(value, string_vars, zero_based_loop_vars);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_zero_based_string_indexes_expr(target, string_vars, zero_based_loop_vars);
            rewrite_zero_based_string_indexes_expr(value, string_vars, zero_based_loop_vars);
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_zero_based_string_indexes_expr(cond, string_vars, zero_based_loop_vars);
            for stmt in then_body {
                rewrite_zero_based_string_indexes_stmt(stmt, string_vars, zero_based_loop_vars);
            }
            for (cond, body) in elifs {
                rewrite_zero_based_string_indexes_expr(cond, string_vars, zero_based_loop_vars);
                for stmt in body {
                    rewrite_zero_based_string_indexes_stmt(stmt, string_vars, zero_based_loop_vars);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_zero_based_string_indexes_stmt(stmt, string_vars, zero_based_loop_vars);
                }
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            rewrite_zero_based_string_indexes_expr(cond, string_vars, zero_based_loop_vars);
            for stmt in body {
                rewrite_zero_based_string_indexes_stmt(stmt, string_vars, zero_based_loop_vars);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_zero_based_string_indexes_stmt(stmt, string_vars, zero_based_loop_vars);
                }
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            for stmt in body {
                rewrite_zero_based_string_indexes_stmt(stmt, string_vars, zero_based_loop_vars);
            }
            rewrite_zero_based_string_indexes_expr(cond, string_vars, zero_based_loop_vars);
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                rewrite_zero_based_string_indexes_stmt(init, string_vars, zero_based_loop_vars);
            }
            if let Some(cond) = cond {
                rewrite_zero_based_string_indexes_expr(cond, string_vars, zero_based_loop_vars);
            }
            if let Some(update) = update {
                rewrite_zero_based_string_indexes_expr(update, string_vars, zero_based_loop_vars);
            }
            let zero_var = init
                .as_ref()
                .and_then(|init| zero_based_for_var(init, cond.as_ref()));
            if let Some(var) = &zero_var {
                zero_based_loop_vars.insert(var.clone());
            }
            for stmt in body {
                rewrite_zero_based_string_indexes_stmt(stmt, string_vars, zero_based_loop_vars);
            }
            if let Some(var) = zero_var {
                zero_based_loop_vars.remove(&var);
            }
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            rewrite_zero_based_string_indexes_expr(iter, string_vars, zero_based_loop_vars);
            for stmt in body {
                rewrite_zero_based_string_indexes_stmt(stmt, string_vars, zero_based_loop_vars);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_zero_based_string_indexes_stmt(stmt, string_vars, zero_based_loop_vars);
                }
            }
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            rewrite_zero_based_string_indexes_expr(expr, string_vars, zero_based_loop_vars);
            for case in cases {
                for stmt in &mut case.body {
                    rewrite_zero_based_string_indexes_stmt(stmt, string_vars, zero_based_loop_vars);
                }
            }
            if let Some(default) = default {
                for stmt in default {
                    rewrite_zero_based_string_indexes_stmt(stmt, string_vars, zero_based_loop_vars);
                }
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            for stmt in body {
                rewrite_zero_based_string_indexes_stmt(stmt, string_vars, zero_based_loop_vars);
            }
            for catch in catches {
                for stmt in &mut catch.body {
                    rewrite_zero_based_string_indexes_stmt(stmt, string_vars, zero_based_loop_vars);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_zero_based_string_indexes_stmt(stmt, string_vars, zero_based_loop_vars);
                }
            }
            if let Some(body) = finally {
                for stmt in body {
                    rewrite_zero_based_string_indexes_stmt(stmt, string_vars, zero_based_loop_vars);
                }
            }
        }
        StmtKind::With { items, body, .. } => {
            for item in items {
                rewrite_zero_based_string_indexes_expr(
                    &mut item.expr,
                    string_vars,
                    zero_based_loop_vars,
                );
            }
            for stmt in body {
                rewrite_zero_based_string_indexes_stmt(stmt, string_vars, zero_based_loop_vars);
            }
        }
        _ => {}
    }
}

fn rewrite_zero_based_string_indexes_member(member: &mut ClassMember) {
    match member {
        ClassMember::Field {
            init: Some(expr), ..
        }
        | ClassMember::Const { value: expr, .. } => {
            let mut strings = std::collections::HashSet::new();
            let mut zero = std::collections::HashSet::new();
            rewrite_zero_based_string_indexes_expr(expr, &mut strings, &mut zero);
        }
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            let mut strings = std::collections::HashSet::new();
            let mut zero = std::collections::HashSet::new();
            rewrite_zero_based_string_indexes_stmt(stmt, &mut strings, &mut zero);
        }
        ClassMember::Constructor { body, .. } => {
            let mut strings = std::collections::HashSet::new();
            let mut zero = std::collections::HashSet::new();
            for stmt in body {
                rewrite_zero_based_string_indexes_stmt(stmt, &mut strings, &mut zero);
            }
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                let mut strings = std::collections::HashSet::new();
                let mut zero = std::collections::HashSet::new();
                for stmt in getter {
                    rewrite_zero_based_string_indexes_stmt(stmt, &mut strings, &mut zero);
                }
            }
            if let Some(setter) = setter {
                let mut strings = std::collections::HashSet::new();
                let mut zero = std::collections::HashSet::new();
                for stmt in &mut setter.body {
                    rewrite_zero_based_string_indexes_stmt(stmt, &mut strings, &mut zero);
                }
            }
        }
        _ => {}
    }
}

fn zero_based_for_var(stmt: &Statement, cond: Option<&Expression>) -> Option<String> {
    match &stmt.kind {
        StmtKind::Assign { targets, value } if targets.len() == 1 && expr_is_int(value, 0) => {
            if let ExprKind::Ident(name) = &targets[0].kind {
                return Some(name.to_lowercase());
            }
        }
        StmtKind::VarDecl { declarations, .. } if declarations.len() == 1 => {
            let decl = &declarations[0];
            if decl.init.as_ref().is_some_and(|expr| expr_is_int(expr, 0)) {
                if let BindingPattern::Ident(name) = &decl.pattern {
                    return Some(name.to_lowercase());
                }
            }
        }
        _ => {}
    }
    let var_name = match &stmt.kind {
        StmtKind::Assign { targets, .. } if targets.len() == 1 => {
            if let ExprKind::Ident(name) = &targets[0].kind {
                Some(name.to_lowercase())
            } else {
                None
            }
        }
        StmtKind::VarDecl { declarations, .. } if declarations.len() == 1 => {
            if let BindingPattern::Ident(name) = &declarations[0].pattern {
                Some(name.to_lowercase())
            } else {
                None
            }
        }
        _ => None,
    };
    if let (Some(name), Some(cond)) = (var_name, cond) {
        if for_condition_ends_at_zero(cond, &name) {
            return Some(name);
        }
    }
    None
}

fn for_condition_ends_at_zero(cond: &Expression, var_name: &str) -> bool {
    match &cond.kind {
        ExprKind::Binary {
            op: BinOp::GtEq,
            left,
            right,
        } if expr_is_int(right, 0) => {
            matches!(&left.kind, ExprKind::Ident(name) if name.eq_ignore_ascii_case(var_name))
        }
        _ => false,
    }
}

fn rewrite_zero_based_string_indexes_expr(
    expr: &mut Expression,
    string_vars: &std::collections::HashSet<String>,
    zero_based_loop_vars: &std::collections::HashSet<String>,
) {
    match &mut expr.kind {
        ExprKind::Index { object, index, .. } => {
            rewrite_zero_based_string_indexes_expr(object, string_vars, zero_based_loop_vars);
            rewrite_zero_based_string_indexes_expr(index, string_vars, zero_based_loop_vars);
            if expr_is_string_receiver(object, string_vars)
                && expr_references_any_var(index, zero_based_loop_vars)
            {
                **index = Expression::new(ExprKind::Binary {
                    op: BinOp::Add,
                    left: Box::new((**index).clone()),
                    right: Box::new(Expression::int(1)),
                });
            }
        }
        ExprKind::Binary { left, right, .. } => {
            rewrite_zero_based_string_indexes_expr(left, string_vars, zero_based_loop_vars);
            rewrite_zero_based_string_indexes_expr(right, string_vars, zero_based_loop_vars);
        }
        ExprKind::Unary { expr: inner, .. }
        | ExprKind::Await(inner)
        | ExprKind::Yield(Some(inner))
        | ExprKind::Spread(inner)
        | ExprKind::TypeOf(inner)
        | ExprKind::Void(inner)
        | ExprKind::Delete(inner)
        | ExprKind::YieldFrom(inner)
        | ExprKind::Cast { expr: inner, .. } => {
            rewrite_zero_based_string_indexes_expr(inner, string_vars, zero_based_loop_vars);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_zero_based_string_indexes_expr(cond, string_vars, zero_based_loop_vars);
            rewrite_zero_based_string_indexes_expr(then, string_vars, zero_based_loop_vars);
            rewrite_zero_based_string_indexes_expr(else_, string_vars, zero_based_loop_vars);
        }
        ExprKind::Call { callee, args, .. } => {
            rewrite_zero_based_string_indexes_expr(callee, string_vars, zero_based_loop_vars);
            for arg in args {
                rewrite_zero_based_string_indexes_expr(
                    &mut arg.value,
                    string_vars,
                    zero_based_loop_vars,
                );
            }
        }
        ExprKind::Member { object, .. } => {
            rewrite_zero_based_string_indexes_expr(object, string_vars, zero_based_loop_vars);
        }
        ExprKind::New { class, args } => {
            rewrite_zero_based_string_indexes_expr(class, string_vars, zero_based_loop_vars);
            for arg in args {
                rewrite_zero_based_string_indexes_expr(
                    &mut arg.value,
                    string_vars,
                    zero_based_loop_vars,
                );
            }
        }
        ExprKind::Array(elements) => {
            for element in elements {
                if let Some(key) = &mut element.key {
                    rewrite_zero_based_string_indexes_expr(key, string_vars, zero_based_loop_vars);
                }
                rewrite_zero_based_string_indexes_expr(
                    &mut element.value,
                    string_vars,
                    zero_based_loop_vars,
                );
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) => {
            for item in items {
                rewrite_zero_based_string_indexes_expr(item, string_vars, zero_based_loop_vars);
            }
        }
        ExprKind::NullCoalesce { left, right } => {
            rewrite_zero_based_string_indexes_expr(left, string_vars, zero_based_loop_vars);
            rewrite_zero_based_string_indexes_expr(right, string_vars, zero_based_loop_vars);
        }
        ExprKind::Range { start, end, .. } => {
            rewrite_zero_based_string_indexes_expr(start, string_vars, zero_based_loop_vars);
            rewrite_zero_based_string_indexes_expr(end, string_vars, zero_based_loop_vars);
        }
        ExprKind::Assign { target, value } | ExprKind::Walrus { target, value } => {
            rewrite_zero_based_string_indexes_expr(target, string_vars, zero_based_loop_vars);
            rewrite_zero_based_string_indexes_expr(value, string_vars, zero_based_loop_vars);
        }
        ExprKind::SuperCall { args, .. } => {
            for arg in args {
                rewrite_zero_based_string_indexes_expr(
                    &mut arg.value,
                    string_vars,
                    zero_based_loop_vars,
                );
            }
        }
        ExprKind::StaticAccess { class, member } => {
            rewrite_zero_based_string_indexes_expr(class, string_vars, zero_based_loop_vars);
            rewrite_zero_based_string_indexes_expr(member, string_vars, zero_based_loop_vars);
        }
        _ => {}
    }
}

fn expr_is_string_receiver(
    expr: &Expression,
    string_vars: &std::collections::HashSet<String>,
) -> bool {
    matches!(&expr.kind, ExprKind::Ident(name) if string_vars.contains(&name.to_lowercase()))
}

fn rewrite_pascal_datetime_arithmetic(body: &mut [Statement]) {
    let mut datetime_vars = std::collections::HashSet::new();
    for stmt in body {
        rewrite_pascal_datetime_arithmetic_stmt(stmt, &mut datetime_vars);
    }
}

fn rewrite_pascal_datetime_arithmetic_stmt(
    stmt: &mut Statement,
    datetime_vars: &mut std::collections::HashSet<String>,
) {
    match &mut stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_pascal_datetime_arithmetic_expr(init, datetime_vars);
                }
                if decl
                    .type_hint
                    .as_deref()
                    .is_some_and(|hint| bare_type_name(hint).eq_ignore_ascii_case("TDateTime"))
                {
                    if let BindingPattern::Ident(name) = &decl.pattern {
                        datetime_vars.insert(name.to_lowercase());
                    }
                }
            }
        }
        StmtKind::FunctionDecl { params, body, .. } => {
            let mut scoped = datetime_vars.clone();
            for param in params {
                if param
                    .type_hint
                    .as_deref()
                    .is_some_and(|hint| bare_type_name(hint).eq_ignore_ascii_case("TDateTime"))
                {
                    scoped.insert(param.name.to_lowercase());
                }
            }
            for stmt in body {
                rewrite_pascal_datetime_arithmetic_stmt(stmt, &mut scoped);
            }
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                rewrite_pascal_datetime_arithmetic_member(member);
            }
        }
        StmtKind::Block(body) | StmtKind::NamespaceDecl { body, .. } => {
            let mut scoped = datetime_vars.clone();
            for stmt in body {
                rewrite_pascal_datetime_arithmetic_stmt(stmt, &mut scoped);
            }
        }
        StmtKind::Expr(expr) => rewrite_pascal_datetime_arithmetic_expr(expr, datetime_vars),
        StmtKind::Return(Some(expr))
        | StmtKind::Throw {
            expr: Some(expr), ..
        } => rewrite_pascal_datetime_arithmetic_expr(expr, datetime_vars),
        StmtKind::Throw {
            cause: Some(cause), ..
        } => rewrite_pascal_datetime_arithmetic_expr(cause, datetime_vars),
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_pascal_datetime_arithmetic_expr(target, datetime_vars);
            }
            rewrite_pascal_datetime_arithmetic_expr(value, datetime_vars);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_pascal_datetime_arithmetic_expr(target, datetime_vars);
            rewrite_pascal_datetime_arithmetic_expr(value, datetime_vars);
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_pascal_datetime_arithmetic_expr(cond, datetime_vars);
            for stmt in then_body {
                rewrite_pascal_datetime_arithmetic_stmt(stmt, datetime_vars);
            }
            for (cond, body) in elifs {
                rewrite_pascal_datetime_arithmetic_expr(cond, datetime_vars);
                for stmt in body {
                    rewrite_pascal_datetime_arithmetic_stmt(stmt, datetime_vars);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_pascal_datetime_arithmetic_stmt(stmt, datetime_vars);
                }
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            rewrite_pascal_datetime_arithmetic_expr(cond, datetime_vars);
            for stmt in body {
                rewrite_pascal_datetime_arithmetic_stmt(stmt, datetime_vars);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_pascal_datetime_arithmetic_stmt(stmt, datetime_vars);
                }
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            for stmt in body {
                rewrite_pascal_datetime_arithmetic_stmt(stmt, datetime_vars);
            }
            rewrite_pascal_datetime_arithmetic_expr(cond, datetime_vars);
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                rewrite_pascal_datetime_arithmetic_stmt(init, datetime_vars);
            }
            if let Some(cond) = cond {
                rewrite_pascal_datetime_arithmetic_expr(cond, datetime_vars);
            }
            if let Some(update) = update {
                rewrite_pascal_datetime_arithmetic_expr(update, datetime_vars);
            }
            for stmt in body {
                rewrite_pascal_datetime_arithmetic_stmt(stmt, datetime_vars);
            }
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            rewrite_pascal_datetime_arithmetic_expr(iter, datetime_vars);
            for stmt in body {
                rewrite_pascal_datetime_arithmetic_stmt(stmt, datetime_vars);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_pascal_datetime_arithmetic_stmt(stmt, datetime_vars);
                }
            }
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            rewrite_pascal_datetime_arithmetic_expr(expr, datetime_vars);
            for case in cases {
                for stmt in &mut case.body {
                    rewrite_pascal_datetime_arithmetic_stmt(stmt, datetime_vars);
                }
            }
            if let Some(default) = default {
                for stmt in default {
                    rewrite_pascal_datetime_arithmetic_stmt(stmt, datetime_vars);
                }
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            for stmt in body {
                rewrite_pascal_datetime_arithmetic_stmt(stmt, datetime_vars);
            }
            for catch in catches {
                for stmt in &mut catch.body {
                    rewrite_pascal_datetime_arithmetic_stmt(stmt, datetime_vars);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_pascal_datetime_arithmetic_stmt(stmt, datetime_vars);
                }
            }
            if let Some(body) = finally {
                for stmt in body {
                    rewrite_pascal_datetime_arithmetic_stmt(stmt, datetime_vars);
                }
            }
        }
        StmtKind::With { items, body, .. } => {
            for item in items {
                rewrite_pascal_datetime_arithmetic_expr(&mut item.expr, datetime_vars);
            }
            for stmt in body {
                rewrite_pascal_datetime_arithmetic_stmt(stmt, datetime_vars);
            }
        }
        _ => {}
    }
}

fn rewrite_pascal_datetime_arithmetic_member(member: &mut ClassMember) {
    match member {
        ClassMember::Field {
            init: Some(expr), ..
        }
        | ClassMember::Const { value: expr, .. } => {
            let mut vars = std::collections::HashSet::new();
            rewrite_pascal_datetime_arithmetic_expr(expr, &mut vars);
        }
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            let mut vars = std::collections::HashSet::new();
            rewrite_pascal_datetime_arithmetic_stmt(stmt, &mut vars);
        }
        ClassMember::Constructor { body, params, .. } => {
            let mut vars = std::collections::HashSet::new();
            for param in params {
                if param
                    .type_hint
                    .as_deref()
                    .is_some_and(|hint| bare_type_name(hint).eq_ignore_ascii_case("TDateTime"))
                {
                    vars.insert(param.name.to_lowercase());
                }
            }
            for stmt in body {
                rewrite_pascal_datetime_arithmetic_stmt(stmt, &mut vars);
            }
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                let mut vars = std::collections::HashSet::new();
                for stmt in getter {
                    rewrite_pascal_datetime_arithmetic_stmt(stmt, &mut vars);
                }
            }
            if let Some(setter) = setter {
                let mut vars = std::collections::HashSet::new();
                for stmt in &mut setter.body {
                    rewrite_pascal_datetime_arithmetic_stmt(stmt, &mut vars);
                }
            }
        }
        _ => {}
    }
}

fn rewrite_pascal_datetime_arithmetic_expr(
    expr: &mut Expression,
    datetime_vars: &std::collections::HashSet<String>,
) {
    match &mut expr.kind {
        ExprKind::Binary { op, left, right } => {
            rewrite_pascal_datetime_arithmetic_expr(left, datetime_vars);
            rewrite_pascal_datetime_arithmetic_expr(right, datetime_vars);
            let left_dt = expr_is_datetime_value(left, datetime_vars);
            let right_dt = expr_is_datetime_value(right, datetime_vars);
            match op {
                BinOp::Sub if left_dt && right_dt => {
                    expr.kind = bin_expr(BinOp::Div, (**left).clone(), (**right).clone()).kind;
                    if let ExprKind::Binary { left, right, .. } = &mut expr.kind {
                        let sub = bin_expr(BinOp::Sub, (**left).clone(), (**right).clone());
                        **left = sub;
                        **right = int_expr(86_400_000);
                    }
                }
                BinOp::Add | BinOp::Sub
                    if left_dt && !right_dt && !expr_is_fixed_datetime_delta(right) =>
                {
                    **right = bin_expr(BinOp::Mul, (**right).clone(), int_expr(86_400_000));
                }
                BinOp::Add if !left_dt && right_dt && !expr_is_fixed_datetime_delta(left) => {
                    **left = bin_expr(BinOp::Mul, (**left).clone(), int_expr(86_400_000));
                }
                _ => {}
            }
        }
        ExprKind::Call { callee, args, .. } => {
            if matches!(&callee.kind, ExprKind::Ident(name) if name == "__pascal_abs") {
                return;
            }
            rewrite_pascal_datetime_arithmetic_expr(callee, datetime_vars);
            for arg in args {
                rewrite_pascal_datetime_arithmetic_expr(&mut arg.value, datetime_vars);
            }
        }
        ExprKind::Unary { expr: inner, .. }
        | ExprKind::Await(inner)
        | ExprKind::Yield(Some(inner))
        | ExprKind::Spread(inner)
        | ExprKind::TypeOf(inner)
        | ExprKind::Void(inner)
        | ExprKind::Delete(inner)
        | ExprKind::YieldFrom(inner)
        | ExprKind::Cast { expr: inner, .. } => {
            rewrite_pascal_datetime_arithmetic_expr(inner, datetime_vars);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_pascal_datetime_arithmetic_expr(cond, datetime_vars);
            rewrite_pascal_datetime_arithmetic_expr(then, datetime_vars);
            rewrite_pascal_datetime_arithmetic_expr(else_, datetime_vars);
        }
        ExprKind::Member { object, .. } => {
            rewrite_pascal_datetime_arithmetic_expr(object, datetime_vars);
        }
        ExprKind::New { class, args } => {
            rewrite_pascal_datetime_arithmetic_expr(class, datetime_vars);
            for arg in args {
                rewrite_pascal_datetime_arithmetic_expr(&mut arg.value, datetime_vars);
            }
        }
        ExprKind::Array(elements) => {
            for element in elements {
                if let Some(key) = &mut element.key {
                    rewrite_pascal_datetime_arithmetic_expr(key, datetime_vars);
                }
                rewrite_pascal_datetime_arithmetic_expr(&mut element.value, datetime_vars);
            }
        }
        ExprKind::Object(properties) => {
            for property in properties {
                match property {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        rewrite_pascal_datetime_arithmetic_expr(key, datetime_vars);
                        rewrite_pascal_datetime_arithmetic_expr(value, datetime_vars);
                    }
                    ObjectProperty::Spread(value) => {
                        rewrite_pascal_datetime_arithmetic_expr(value, datetime_vars);
                    }
                    _ => {}
                }
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => {
            for item in items {
                rewrite_pascal_datetime_arithmetic_expr(item, datetime_vars);
            }
        }
        ExprKind::NullCoalesce { left, right } => {
            rewrite_pascal_datetime_arithmetic_expr(left, datetime_vars);
            rewrite_pascal_datetime_arithmetic_expr(right, datetime_vars);
        }
        ExprKind::Range { start, end, .. } => {
            rewrite_pascal_datetime_arithmetic_expr(start, datetime_vars);
            rewrite_pascal_datetime_arithmetic_expr(end, datetime_vars);
        }
        ExprKind::Assign { target, value } | ExprKind::Walrus { target, value } => {
            rewrite_pascal_datetime_arithmetic_expr(target, datetime_vars);
            rewrite_pascal_datetime_arithmetic_expr(value, datetime_vars);
        }
        ExprKind::SuperCall { args, .. } => {
            for arg in args {
                rewrite_pascal_datetime_arithmetic_expr(&mut arg.value, datetime_vars);
            }
        }
        ExprKind::StaticAccess { class, member } => {
            rewrite_pascal_datetime_arithmetic_expr(class, datetime_vars);
            rewrite_pascal_datetime_arithmetic_expr(member, datetime_vars);
        }
        _ => {}
    }
}

fn expr_is_datetime_value(
    expr: &Expression,
    datetime_vars: &std::collections::HashSet<String>,
) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => datetime_vars.contains(&name.to_lowercase()),
        ExprKind::Call { callee, .. } => {
            matches!(&callee.kind, ExprKind::Ident(name) if name == "__pascal_date_utc")
        }
        _ => false,
    }
}

fn expr_is_fixed_datetime_delta(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Binary {
            op: BinOp::Mul,
            left,
            right,
        } => expr_is_time_unit_ms(left) || expr_is_time_unit_ms(right),
        _ => false,
    }
}

fn expr_is_time_unit_ms(expr: &Expression) -> bool {
    matches!(
        expr.kind,
        ExprKind::Lit(Literal::Int(60_000 | 3_600_000 | 86_400_000))
    )
}

fn expr_references_any_var(expr: &Expression, names: &std::collections::HashSet<String>) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => names.contains(&name.to_lowercase()),
        ExprKind::Binary { left, right, .. } => {
            expr_references_any_var(left, names) || expr_references_any_var(right, names)
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Await(expr)
        | ExprKind::Yield(Some(expr))
        | ExprKind::Spread(expr)
        | ExprKind::TypeOf(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Cast { expr, .. } => expr_references_any_var(expr, names),
        ExprKind::Ternary { cond, then, else_ } => {
            expr_references_any_var(cond, names)
                || expr_references_any_var(then, names)
                || expr_references_any_var(else_, names)
        }
        ExprKind::Call { callee, args, .. } => {
            expr_references_any_var(callee, names)
                || args
                    .iter()
                    .any(|arg| expr_references_any_var(&arg.value, names))
        }
        ExprKind::Member { object, .. } => expr_references_any_var(object, names),
        ExprKind::Index { object, index, .. } => {
            expr_references_any_var(object, names) || expr_references_any_var(index, names)
        }
        ExprKind::New { class, args } => {
            expr_references_any_var(class, names)
                || args
                    .iter()
                    .any(|arg| expr_references_any_var(&arg.value, names))
        }
        ExprKind::Array(elements) => elements.iter().any(|element| {
            element
                .key
                .as_ref()
                .is_some_and(|key| expr_references_any_var(key, names))
                || expr_references_any_var(&element.value, names)
        }),
        ExprKind::Tuple(items) | ExprKind::Set(items) => items
            .iter()
            .any(|item| expr_references_any_var(item, names)),
        ExprKind::NullCoalesce { left, right } => {
            expr_references_any_var(left, names) || expr_references_any_var(right, names)
        }
        ExprKind::Range { start, end, .. } => {
            expr_references_any_var(start, names) || expr_references_any_var(end, names)
        }
        ExprKind::Assign { target, value } | ExprKind::Walrus { target, value } => {
            expr_references_any_var(target, names) || expr_references_any_var(value, names)
        }
        ExprKind::SuperCall { args, .. } => args
            .iter()
            .any(|arg| expr_references_any_var(&arg.value, names)),
        ExprKind::StaticAccess { class, member } => {
            expr_references_any_var(class, names) || expr_references_any_var(member, names)
        }
        _ => false,
    }
}

fn expr_is_int(expr: &Expression, expected: i64) -> bool {
    matches!(expr.kind, ExprKind::Lit(Literal::Int(value)) if value == expected)
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
                                struct_fields.get(&type_name.to_lowercase()).map(|fields| {
                                    build_struct_copy_statements(target, source, type_name, fields)
                                })
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
                lower_struct_copy_assignments_in_block(
                    &mut catch.body,
                    struct_fields,
                    &mut catch_scope,
                );
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

fn lower_pascal_array_value_semantics(body: &mut [Statement]) {
    let mut env = std::collections::HashMap::new();
    lower_pascal_array_value_semantics_block(body, &mut env);
}

fn lower_pascal_array_value_semantics_block(
    body: &mut [Statement],
    env: &mut std::collections::HashMap<String, String>,
) {
    for stmt in body {
        lower_pascal_array_value_semantics_stmt(stmt, env);
    }
}

fn lower_pascal_array_value_semantics_stmt(
    stmt: &mut Statement,
    env: &mut std::collections::HashMap<String, String>,
) {
    match &mut stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations.iter_mut() {
                if let Some(init) = &mut decl.init {
                    lower_pascal_array_value_semantics_expr(init, env);
                }
            }
            for decl in declarations {
                if let (BindingPattern::Ident(name), Some(type_hint)) =
                    (&decl.pattern, &decl.type_hint)
                {
                    env.insert(name.to_lowercase(), type_hint.to_string());
                }
            }
        }
        StmtKind::Assign { targets, value } if targets.len() == 1 => {
            lower_pascal_array_value_semantics_expr(value, env);
            if pascal_expr_is_array_like(value, env) && pascal_expr_is_array_like(&targets[0], env)
            {
                *value = pascal_array_clone_expr(value.clone());
            }
            lower_pascal_array_value_semantics_expr(&mut targets[0], env);
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                lower_pascal_array_value_semantics_expr(target, env);
            }
            lower_pascal_array_value_semantics_expr(value, env);
        }
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            lower_pascal_array_value_semantics_expr(expr, env);
        }
        StmtKind::Return(None) | StmtKind::Empty => {}
        StmtKind::Block(inner) => {
            let mut scoped = env.clone();
            lower_pascal_array_value_semantics_block(inner, &mut scoped);
        }
        StmtKind::FunctionDecl {
            params,
            return_type,
            body,
            ..
        } => {
            let mut scoped = std::collections::HashMap::new();
            for param in params {
                if let Some(type_hint) = &param.type_hint {
                    scoped.insert(param.name.to_lowercase(), type_hint.to_string());
                }
            }
            if let Some(return_type) = return_type {
                if is_pascal_array_type_hint(return_type) {
                    scoped.insert("result".to_string(), return_type.to_string());
                }
            }
            lower_pascal_array_value_semantics_block(body, &mut scoped);
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            lower_pascal_array_value_semantics_expr(cond, env);
            let mut then_scope = env.clone();
            lower_pascal_array_value_semantics_block(then_body, &mut then_scope);
            for (cond, body) in elifs {
                lower_pascal_array_value_semantics_expr(cond, env);
                let mut elif_scope = env.clone();
                lower_pascal_array_value_semantics_block(body, &mut elif_scope);
            }
            if let Some(body) = else_body {
                let mut else_scope = env.clone();
                lower_pascal_array_value_semantics_block(body, &mut else_scope);
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            let mut scoped = env.clone();
            if let Some(init) = init {
                lower_pascal_array_value_semantics_stmt(init, &mut scoped);
            }
            if let Some(cond) = cond {
                lower_pascal_array_value_semantics_expr(cond, &scoped);
            }
            if let Some(update) = update {
                lower_pascal_array_value_semantics_expr(update, &scoped);
            }
            lower_pascal_array_value_semantics_block(body, &mut scoped);
        }
        StmtKind::ForIn { iter, body, .. } => {
            lower_pascal_array_value_semantics_expr(iter, env);
            let mut scoped = env.clone();
            lower_pascal_array_value_semantics_block(body, &mut scoped);
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            lower_pascal_array_value_semantics_expr(cond, env);
            let mut scoped = env.clone();
            lower_pascal_array_value_semantics_block(body, &mut scoped);
            if let Some(body) = else_body {
                let mut else_scope = env.clone();
                lower_pascal_array_value_semantics_block(body, &mut else_scope);
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            let mut scoped = env.clone();
            lower_pascal_array_value_semantics_block(body, &mut scoped);
            lower_pascal_array_value_semantics_expr(cond, env);
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            let mut try_scope = env.clone();
            lower_pascal_array_value_semantics_block(body, &mut try_scope);
            for catch in catches {
                let mut catch_scope = env.clone();
                lower_pascal_array_value_semantics_block(&mut catch.body, &mut catch_scope);
            }
            if let Some(body) = else_body {
                let mut else_scope = env.clone();
                lower_pascal_array_value_semantics_block(body, &mut else_scope);
            }
            if let Some(body) = finally {
                let mut finally_scope = env.clone();
                lower_pascal_array_value_semantics_block(body, &mut finally_scope);
            }
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                lower_pascal_array_value_semantics_member(member, env);
            }
        }
        _ => {}
    }
}

fn lower_pascal_array_value_semantics_member(
    member: &mut ClassMember,
    env: &std::collections::HashMap<String, String>,
) {
    match member {
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            let mut scoped = env.clone();
            lower_pascal_array_value_semantics_stmt(stmt, &mut scoped);
        }
        ClassMember::Constructor { params, body, .. } => {
            let mut scoped = env.clone();
            for param in params {
                if let Some(type_hint) = &param.type_hint {
                    scoped.insert(param.name.to_lowercase(), type_hint.to_string());
                }
            }
            lower_pascal_array_value_semantics_block(body, &mut scoped);
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                let mut scoped = env.clone();
                lower_pascal_array_value_semantics_block(getter, &mut scoped);
            }
            if let Some(setter) = setter {
                let mut scoped = env.clone();
                if let Some(type_hint) = &setter.param.type_hint {
                    if is_pascal_array_type_hint(type_hint) {
                        scoped.insert(setter.param.name.to_lowercase(), type_hint.to_string());
                    }
                }
                lower_pascal_array_value_semantics_block(&mut setter.body, &mut scoped);
            }
        }
        _ => {}
    }
}

fn lower_pascal_array_value_semantics_expr(
    expr: &mut Expression,
    env: &std::collections::HashMap<String, String>,
) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            lower_pascal_array_value_semantics_expr(callee, env);
            for arg in args.iter_mut() {
                lower_pascal_array_value_semantics_expr(&mut arg.value, env);
            }
            if let ExprKind::Ident(name) = &callee.kind {
                if name.eq_ignore_ascii_case("copy") && args.len() >= 2 {
                    let source = args[0].value.clone();
                    if pascal_expr_is_array_like(&source, env) {
                        let start = args[1].value.clone();
                        let count = args
                            .get(2)
                            .map(|arg| arg.value.clone())
                            .unwrap_or_else(|| pascal_array_length_expr(source.clone()));
                        *expr = pascal_array_copy_expr(source, start, count);
                    }
                }
            }
        }
        ExprKind::Binary { op, left, right } => {
            lower_pascal_array_value_semantics_expr(left, env);
            lower_pascal_array_value_semantics_expr(right, env);
            if matches!(op, BinOp::Add)
                && (pascal_expr_is_array_like(left, env) || pascal_expr_is_array_like(right, env))
            {
                *expr = pascal_array_concat_expr((**left).clone(), (**right).clone());
            } else if matches!(op, BinOp::BitXor)
                && pascal_expr_is_pascal_integer_like(left, env)
                && pascal_expr_is_pascal_integer_like(right, env)
            {
                *expr = pascal_call(
                    "__pascal_int_xor",
                    vec![(**left).clone(), (**right).clone()],
                );
            }
        }
        ExprKind::Member { object, .. } => lower_pascal_array_value_semantics_expr(object, env),
        ExprKind::Index { object, index, .. } => {
            lower_pascal_array_value_semantics_expr(object, env);
            lower_pascal_array_value_semantics_expr(index, env);
        }
        ExprKind::Unary { expr, .. } => lower_pascal_array_value_semantics_expr(expr, env),
        ExprKind::Ternary { cond, then, else_ } => {
            lower_pascal_array_value_semantics_expr(cond, env);
            lower_pascal_array_value_semantics_expr(then, env);
            lower_pascal_array_value_semantics_expr(else_, env);
        }
        ExprKind::Assign { target, value } => {
            lower_pascal_array_value_semantics_expr(value, env);
            if pascal_expr_is_array_like(value, env) && pascal_expr_is_array_like(target, env) {
                **value = pascal_array_clone_expr((**value).clone());
            }
            lower_pascal_array_value_semantics_expr(target, env);
        }
        ExprKind::Array(items) => {
            for item in items {
                lower_pascal_array_value_semantics_expr(&mut item.value, env);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                if let ObjectProperty::KeyValue { key, value } = prop {
                    lower_pascal_array_value_semantics_expr(key, env);
                    lower_pascal_array_value_semantics_expr(value, env);
                }
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) => {
            for item in items {
                lower_pascal_array_value_semantics_expr(item, env);
            }
        }
        ExprKind::New { class, args } => {
            lower_pascal_array_value_semantics_expr(class, env);
            for arg in args {
                lower_pascal_array_value_semantics_expr(&mut arg.value, env);
            }
        }
        _ => {}
    }
}

fn is_pascal_array_type_hint(type_hint: &str) -> bool {
    type_hint
        .trim_start()
        .get(..5)
        .is_some_and(|prefix| prefix.eq_ignore_ascii_case("array"))
}

fn pascal_expr_is_array_like(
    expr: &Expression,
    env: &std::collections::HashMap<String, String>,
) -> bool {
    match &expr.kind {
        ExprKind::Array(_) => true,
        ExprKind::Ident(name) => env
            .get(&name.to_lowercase())
            .is_some_and(|hint| is_pascal_array_type_hint(hint)),
        ExprKind::Call { callee, .. } => matches!(
            &callee.kind,
            ExprKind::Ident(name)
                if name.eq_ignore_ascii_case("__pascal_array_slice")
                    || name.eq_ignore_ascii_case("__pascal_array_concat")
        ),
        _ => false,
    }
}

fn pascal_expr_is_pascal_integer_like(
    expr: &Expression,
    env: &std::collections::HashMap<String, String>,
) -> bool {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(_)) => true,
        ExprKind::Ident(name) => env
            .get(&name.to_lowercase())
            .is_some_and(|hint| pascal_type_hint_is_integer(hint)),
        ExprKind::Index { object, .. } => pascal_array_index_is_integer(object, env),
        ExprKind::Cast { type_name, .. } => pascal_type_hint_is_integer(type_name),
        _ => false,
    }
}

fn pascal_array_index_is_integer(
    object: &Expression,
    env: &std::collections::HashMap<String, String>,
) -> bool {
    match &object.kind {
        ExprKind::Ident(name) => env
            .get(&name.to_lowercase())
            .and_then(|hint| pascal_array_element_type(hint))
            .is_some_and(|element| pascal_type_hint_is_integer(&element)),
        ExprKind::Index { object, .. } => pascal_array_index_is_integer(object, env),
        _ => false,
    }
}

fn pascal_type_hint_is_integer(type_hint: &str) -> bool {
    matches!(
        bare_type_name(type_hint).to_ascii_lowercase().as_str(),
        "integer"
            | "int"
            | "longint"
            | "shortint"
            | "smallint"
            | "byte"
            | "word"
            | "cardinal"
            | "int64"
            | "uint64"
            | "longword"
    )
}

fn pascal_array_copy_expr(source: Expression, start: Expression, count: Expression) -> Expression {
    let start0 = Expression::new(ExprKind::Binary {
        op: BinOp::Sub,
        left: Box::new(start),
        right: Box::new(Expression::new(ExprKind::Lit(Literal::Int(1)))),
    });
    let end = Expression::new(ExprKind::Binary {
        op: BinOp::Add,
        left: Box::new(start0.clone()),
        right: Box::new(count),
    });
    pascal_call("__pascal_array_slice", vec![source, start0, end])
}

fn pascal_array_clone_expr(source: Expression) -> Expression {
    pascal_call(
        "__pascal_array_slice",
        vec![
            source.clone(),
            Expression::new(ExprKind::Lit(Literal::Int(0))),
            pascal_array_length_expr(source),
        ],
    )
}

fn pascal_array_concat_expr(left: Expression, right: Expression) -> Expression {
    pascal_call("__pascal_array_concat", vec![left, right])
}

fn pascal_array_length_expr(source: Expression) -> Expression {
    pascal_call("__len__", vec![source])
}

fn rewrite_pascal_fixed_array_bounds(body: &mut [Statement]) {
    let mut env = std::collections::HashMap::new();
    rewrite_pascal_fixed_array_bounds_block(body, &mut env);
}

fn rewrite_pascal_fixed_array_bounds_block(
    body: &mut [Statement],
    env: &mut std::collections::HashMap<String, Vec<(i64, i64)>>,
) {
    for stmt in body {
        rewrite_pascal_fixed_array_bounds_stmt(stmt, env);
    }
}

fn rewrite_pascal_fixed_array_bounds_stmt(
    stmt: &mut Statement,
    env: &mut std::collections::HashMap<String, Vec<(i64, i64)>>,
) {
    match &mut stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations.iter_mut() {
                if let Some(init) = &mut decl.init {
                    rewrite_pascal_fixed_array_bounds_expr(init, env);
                }
            }
            for decl in declarations {
                if let (BindingPattern::Ident(name), Some(type_hint)) =
                    (&decl.pattern, &decl.type_hint)
                {
                    if let Some(bounds) = pascal_const_array_bounds(type_hint) {
                        env.insert(name.to_lowercase(), bounds);
                    }
                }
            }
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_pascal_fixed_array_bounds_expr(target, env);
            }
            rewrite_pascal_fixed_array_bounds_expr(value, env);
        }
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            rewrite_pascal_fixed_array_bounds_expr(expr, env);
        }
        StmtKind::Block(inner) => {
            let mut scoped = env.clone();
            rewrite_pascal_fixed_array_bounds_block(inner, &mut scoped);
        }
        StmtKind::FunctionDecl { params, body, .. } => {
            let mut scoped = std::collections::HashMap::new();
            for param in params {
                if let Some(type_hint) = &param.type_hint {
                    if let Some(bounds) = pascal_const_array_bounds(type_hint) {
                        scoped.insert(param.name.to_lowercase(), bounds);
                    }
                }
            }
            rewrite_pascal_fixed_array_bounds_block(body, &mut scoped);
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_pascal_fixed_array_bounds_expr(cond, env);
            let mut then_scope = env.clone();
            rewrite_pascal_fixed_array_bounds_block(then_body, &mut then_scope);
            for (cond, body) in elifs {
                rewrite_pascal_fixed_array_bounds_expr(cond, env);
                let mut elif_scope = env.clone();
                rewrite_pascal_fixed_array_bounds_block(body, &mut elif_scope);
            }
            if let Some(body) = else_body {
                let mut else_scope = env.clone();
                rewrite_pascal_fixed_array_bounds_block(body, &mut else_scope);
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            let mut scoped = env.clone();
            if let Some(init) = init {
                rewrite_pascal_fixed_array_bounds_stmt(init, &mut scoped);
            }
            if let Some(cond) = cond {
                rewrite_pascal_fixed_array_bounds_expr(cond, &scoped);
            }
            if let Some(update) = update {
                rewrite_pascal_fixed_array_bounds_expr(update, &scoped);
            }
            rewrite_pascal_fixed_array_bounds_block(body, &mut scoped);
        }
        StmtKind::ForIn { iter, body, .. } => {
            rewrite_pascal_fixed_array_bounds_expr(iter, env);
            let mut scoped = env.clone();
            rewrite_pascal_fixed_array_bounds_block(body, &mut scoped);
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            rewrite_pascal_fixed_array_bounds_expr(cond, env);
            let mut scoped = env.clone();
            rewrite_pascal_fixed_array_bounds_block(body, &mut scoped);
            if let Some(body) = else_body {
                let mut else_scope = env.clone();
                rewrite_pascal_fixed_array_bounds_block(body, &mut else_scope);
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            let mut scoped = env.clone();
            rewrite_pascal_fixed_array_bounds_block(body, &mut scoped);
            rewrite_pascal_fixed_array_bounds_expr(cond, env);
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            let mut try_scope = env.clone();
            rewrite_pascal_fixed_array_bounds_block(body, &mut try_scope);
            for catch in catches {
                let mut catch_scope = env.clone();
                rewrite_pascal_fixed_array_bounds_block(&mut catch.body, &mut catch_scope);
            }
            if let Some(body) = else_body {
                let mut else_scope = env.clone();
                rewrite_pascal_fixed_array_bounds_block(body, &mut else_scope);
            }
            if let Some(body) = finally {
                let mut finally_scope = env.clone();
                rewrite_pascal_fixed_array_bounds_block(body, &mut finally_scope);
            }
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                rewrite_pascal_fixed_array_bounds_member(member, env);
            }
        }
        _ => {}
    }
}

fn rewrite_pascal_fixed_array_bounds_member(
    member: &mut ClassMember,
    env: &std::collections::HashMap<String, Vec<(i64, i64)>>,
) {
    match member {
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            let mut scoped = env.clone();
            rewrite_pascal_fixed_array_bounds_stmt(stmt, &mut scoped);
        }
        ClassMember::Constructor { params, body, .. } => {
            let mut scoped = env.clone();
            for param in params {
                if let Some(type_hint) = &param.type_hint {
                    if let Some(bounds) = pascal_const_array_bounds(type_hint) {
                        scoped.insert(param.name.to_lowercase(), bounds);
                    }
                }
            }
            rewrite_pascal_fixed_array_bounds_block(body, &mut scoped);
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                let mut scoped = env.clone();
                rewrite_pascal_fixed_array_bounds_block(getter, &mut scoped);
            }
            if let Some(setter) = setter {
                let mut scoped = env.clone();
                rewrite_pascal_fixed_array_bounds_block(&mut setter.body, &mut scoped);
            }
        }
        _ => {}
    }
}

fn rewrite_pascal_fixed_array_bounds_expr(
    expr: &mut Expression,
    env: &std::collections::HashMap<String, Vec<(i64, i64)>>,
) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            rewrite_pascal_fixed_array_bounds_expr(callee, env);
            for arg in args.iter_mut() {
                rewrite_pascal_fixed_array_bounds_expr(&mut arg.value, env);
            }
            let ExprKind::Ident(name) = &callee.kind else {
                return;
            };
            if !(name.eq_ignore_ascii_case("low") || name.eq_ignore_ascii_case("high")) {
                return;
            }
            let Some(first) = args.first() else {
                return;
            };
            let Some(bounds) = pascal_bounds_for_expr(&first.value, env) else {
                return;
            };
            let dim = args
                .get(1)
                .and_then(|arg| const_int_expr(&arg.value))
                .unwrap_or(1);
            if dim < 1 {
                return;
            }
            let Some((lo, hi)) = bounds.get((dim - 1) as usize).copied() else {
                return;
            };
            *expr = Expression::new(ExprKind::Lit(Literal::Int(if name.eq_ignore_ascii_case("low") {
                lo
            } else {
                hi
            })));
        }
        ExprKind::Binary { left, right, .. } => {
            rewrite_pascal_fixed_array_bounds_expr(left, env);
            rewrite_pascal_fixed_array_bounds_expr(right, env);
        }
        ExprKind::Member { object, .. } => rewrite_pascal_fixed_array_bounds_expr(object, env),
        ExprKind::Index { object, index, .. } => {
            rewrite_pascal_fixed_array_bounds_expr(object, env);
            rewrite_pascal_fixed_array_bounds_expr(index, env);
        }
        ExprKind::Unary { expr, .. } => rewrite_pascal_fixed_array_bounds_expr(expr, env),
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_pascal_fixed_array_bounds_expr(cond, env);
            rewrite_pascal_fixed_array_bounds_expr(then, env);
            rewrite_pascal_fixed_array_bounds_expr(else_, env);
        }
        ExprKind::Assign { target, value } => {
            rewrite_pascal_fixed_array_bounds_expr(target, env);
            rewrite_pascal_fixed_array_bounds_expr(value, env);
        }
        ExprKind::Array(items) => {
            for item in items {
                rewrite_pascal_fixed_array_bounds_expr(&mut item.value, env);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                if let ObjectProperty::KeyValue { key, value } = prop {
                    rewrite_pascal_fixed_array_bounds_expr(key, env);
                    rewrite_pascal_fixed_array_bounds_expr(value, env);
                }
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) => {
            for item in items {
                rewrite_pascal_fixed_array_bounds_expr(item, env);
            }
        }
        ExprKind::New { class, args } => {
            rewrite_pascal_fixed_array_bounds_expr(class, env);
            for arg in args {
                rewrite_pascal_fixed_array_bounds_expr(&mut arg.value, env);
            }
        }
        _ => {}
    }
}

fn pascal_bounds_for_expr(
    expr: &Expression,
    env: &std::collections::HashMap<String, Vec<(i64, i64)>>,
) -> Option<Vec<(i64, i64)>> {
    match &expr.kind {
        ExprKind::Ident(name) => env.get(&name.to_lowercase()).cloned(),
        _ => None,
    }
}

fn pascal_const_array_bounds(type_hint: &str) -> Option<Vec<(i64, i64)>> {
    let trimmed = type_hint.trim();
    let lower = trimmed.to_ascii_lowercase();
    let rest = lower.strip_prefix("array[")?;
    let close = rest.find(']')?;
    let bounds = &trimmed["array[".len().."array[".len() + close];
    let mut out = Vec::new();
    for dim in bounds.split(',') {
        let (lo, hi) = dim.split_once("..")?;
        let lo = lo.trim().parse::<i64>().ok()?;
        let hi = hi.trim().parse::<i64>().ok()?;
        out.push((lo, hi));
    }
    (!out.is_empty()).then_some(out)
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
        StmtKind::StructDecl { name, members, .. } => {
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

fn collect_enum_type_names(body: &[Statement]) -> std::collections::HashSet<String> {
    let mut out = std::collections::HashSet::new();
    collect_enum_metadata(body, &mut out, &mut std::collections::HashMap::new(), None);
    out
}

fn collect_enum_type_counts(body: &[Statement]) -> std::collections::HashMap<String, usize> {
    let mut out = std::collections::HashMap::new();
    collect_enum_metadata(body, &mut std::collections::HashSet::new(), &mut out, None);
    out
}

fn collect_enum_member_ordinals(body: &[Statement]) -> std::collections::HashMap<String, i64> {
    let mut out = std::collections::HashMap::new();
    collect_enum_metadata(
        body,
        &mut std::collections::HashSet::new(),
        &mut std::collections::HashMap::new(),
        Some(&mut out),
    );
    out
}

fn collect_enum_metadata(
    body: &[Statement],
    names: &mut std::collections::HashSet<String>,
    counts: &mut std::collections::HashMap<String, usize>,
    mut ordinals: Option<&mut std::collections::HashMap<String, i64>>,
) {
    for stmt in body {
        match &stmt.kind {
            StmtKind::EnumDecl { name, members, .. } => {
                names.insert(name.to_lowercase());
                counts.insert(name.to_lowercase(), members.len());
                if let Some(ordinals) = ordinals.as_deref_mut() {
                    let mut next = 0i64;
                    for member in members {
                        if let Some(value) = &member.value {
                            if let Some(int_value) = const_int_expr(value) {
                                next = int_value;
                            }
                        }
                        ordinals.insert(member.name.clone(), next);
                        next += 1;
                    }
                }
            }
            StmtKind::FunctionDecl { body, .. } | StmtKind::Block(body) => {
                collect_enum_metadata(body, names, counts, ordinals.as_deref_mut());
            }
            StmtKind::ClassDecl { members, .. }
            | StmtKind::StructDecl { members, .. }
            | StmtKind::ModuleDecl { members, .. } => {
                for member in members {
                    match member {
                        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
                            collect_enum_metadata(
                                std::slice::from_ref(stmt),
                                names,
                                counts,
                                ordinals.as_deref_mut(),
                            );
                        }
                        ClassMember::Constructor { body, .. } => {
                            collect_enum_metadata(body, names, counts, ordinals.as_deref_mut());
                        }
                        ClassMember::Property { getter, setter, .. } => {
                            if let Some(getter) = getter {
                                collect_enum_metadata(
                                    getter,
                                    names,
                                    counts,
                                    ordinals.as_deref_mut(),
                                );
                            }
                            if let Some(setter) = setter {
                                collect_enum_metadata(
                                    &setter.body,
                                    names,
                                    counts,
                                    ordinals.as_deref_mut(),
                                );
                            }
                        }
                        _ => {}
                    }
                }
            }
            _ => {}
        }
    }
}

fn rename_shadowing_pascal_set_vars(
    body: &mut [Statement],
    enum_member_ordinals: &std::collections::HashMap<String, i64>,
) {
    let enum_members = enum_member_ordinals
        .keys()
        .map(|name| name.to_lowercase())
        .collect::<std::collections::HashSet<_>>();
    if enum_members.is_empty() {
        return;
    }
    let mut renames = std::collections::HashMap::new();
    let mut counter = 0usize;
    for stmt in body.iter_mut() {
        collect_shadowing_pascal_set_var_renames_stmt(
            stmt,
            &enum_members,
            &mut renames,
            &mut counter,
        );
    }
    if renames.is_empty() {
        return;
    }
    for stmt in body {
        apply_pascal_ident_renames_stmt(stmt, &renames);
    }
}

fn collect_shadowing_pascal_set_var_renames_stmt(
    stmt: &mut Statement,
    enum_members: &std::collections::HashSet<String>,
    renames: &mut std::collections::HashMap<String, String>,
    counter: &mut usize,
) {
    match &mut stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if !decl
                    .type_hint
                    .as_deref()
                    .is_some_and(is_pascal_set_type_hint)
                {
                    continue;
                }
                let BindingPattern::Ident(name) = &decl.pattern else {
                    continue;
                };
                if enum_members.contains(&name.to_lowercase()) && !renames.contains_key(name) {
                    let safe = format!("__pascal_set_{}_{}", name.to_lowercase(), *counter);
                    *counter += 1;
                    renames.insert(name.clone(), safe);
                }
            }
        }
        StmtKind::FunctionDecl { params, body, .. } => {
            for param in params {
                if param
                    .type_hint
                    .as_deref()
                    .is_some_and(is_pascal_set_type_hint)
                    && enum_members.contains(&param.name.to_lowercase())
                    && !renames.contains_key(&param.name)
                {
                    let safe = format!("__pascal_set_{}_{}", param.name.to_lowercase(), *counter);
                    *counter += 1;
                    renames.insert(param.name.clone(), safe);
                }
            }
            for stmt in body {
                collect_shadowing_pascal_set_var_renames_stmt(stmt, enum_members, renames, counter);
            }
        }
        StmtKind::Block(body) | StmtKind::NamespaceDecl { body, .. } => {
            for stmt in body {
                collect_shadowing_pascal_set_var_renames_stmt(stmt, enum_members, renames, counter);
            }
        }
        _ => {}
    }
}

fn apply_pascal_ident_renames_stmt(
    stmt: &mut Statement,
    renames: &std::collections::HashMap<String, String>,
) {
    match &mut stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let BindingPattern::Ident(name) = &mut decl.pattern {
                    if let Some(replacement) = renames.get(name) {
                        *name = replacement.clone();
                    }
                }
                if let Some(init) = &mut decl.init {
                    apply_pascal_ident_renames_expr(init, renames);
                }
            }
        }
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            apply_pascal_ident_renames_expr(expr, renames)
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                apply_pascal_ident_renames_expr(target, renames);
            }
            apply_pascal_ident_renames_expr(value, renames);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            apply_pascal_ident_renames_expr(target, renames);
            apply_pascal_ident_renames_expr(value, renames);
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            apply_pascal_ident_renames_expr(cond, renames);
            for stmt in then_body {
                apply_pascal_ident_renames_stmt(stmt, renames);
            }
            for (cond, body) in elifs {
                apply_pascal_ident_renames_expr(cond, renames);
                for stmt in body {
                    apply_pascal_ident_renames_stmt(stmt, renames);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    apply_pascal_ident_renames_stmt(stmt, renames);
                }
            }
        }
        StmtKind::While { cond, body, .. } | StmtKind::DoWhile { cond, body, .. } => {
            apply_pascal_ident_renames_expr(cond, renames);
            for stmt in body {
                apply_pascal_ident_renames_stmt(stmt, renames);
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                apply_pascal_ident_renames_stmt(init, renames);
            }
            if let Some(cond) = cond {
                apply_pascal_ident_renames_expr(cond, renames);
            }
            if let Some(update) = update {
                apply_pascal_ident_renames_expr(update, renames);
            }
            for stmt in body {
                apply_pascal_ident_renames_stmt(stmt, renames);
            }
        }
        StmtKind::ForIn {
            var, iter, body, ..
        } => {
            if let Some(replacement) = renames.get(var) {
                *var = replacement.clone();
            }
            apply_pascal_ident_renames_expr(iter, renames);
            for stmt in body {
                apply_pascal_ident_renames_stmt(stmt, renames);
            }
        }
        StmtKind::FunctionDecl { params, body, .. } => {
            for param in params {
                if let Some(replacement) = renames.get(&param.name) {
                    param.name = replacement.clone();
                }
                if let Some(default) = &mut param.default {
                    apply_pascal_ident_renames_expr(default, renames);
                }
            }
            for stmt in body {
                apply_pascal_ident_renames_stmt(stmt, renames);
            }
        }
        StmtKind::Block(body) | StmtKind::NamespaceDecl { body, .. } => {
            for stmt in body {
                apply_pascal_ident_renames_stmt(stmt, renames);
            }
        }
        _ => {}
    }
}

fn apply_pascal_ident_renames_expr(
    expr: &mut Expression,
    renames: &std::collections::HashMap<String, String>,
) {
    match &mut expr.kind {
        ExprKind::Ident(name) => {
            if let Some(replacement) = renames.get(name) {
                *name = replacement.clone();
            }
        }
        ExprKind::Call { callee, args, .. } => {
            apply_pascal_ident_renames_expr(callee, renames);
            for arg in args {
                apply_pascal_ident_renames_expr(&mut arg.value, renames);
            }
        }
        ExprKind::Member { object, .. } => apply_pascal_ident_renames_expr(object, renames),
        ExprKind::Index { object, index, .. } => {
            apply_pascal_ident_renames_expr(object, renames);
            apply_pascal_ident_renames_expr(index, renames);
        }
        ExprKind::Unary { expr, .. } => apply_pascal_ident_renames_expr(expr, renames),
        ExprKind::Binary { left, right, .. } => {
            apply_pascal_ident_renames_expr(left, renames);
            apply_pascal_ident_renames_expr(right, renames);
        }
        ExprKind::Assign { target, value } | ExprKind::Walrus { target, value } => {
            apply_pascal_ident_renames_expr(target, renames);
            apply_pascal_ident_renames_expr(value, renames);
        }
        ExprKind::Array(elements) => {
            for element in elements {
                apply_pascal_ident_renames_expr(&mut element.value, renames);
                if let Some(key) = &mut element.key {
                    apply_pascal_ident_renames_expr(key, renames);
                }
            }
        }
        ExprKind::Set(items) | ExprKind::Tuple(items) | ExprKind::Sequence(items) => {
            for item in items {
                apply_pascal_ident_renames_expr(item, renames);
            }
        }
        ExprKind::Object(properties) => {
            for property in properties {
                match property {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        apply_pascal_ident_renames_expr(key, renames);
                        apply_pascal_ident_renames_expr(value, renames);
                    }
                    ObjectProperty::Spread(value) => {
                        apply_pascal_ident_renames_expr(value, renames)
                    }
                    _ => {}
                }
            }
        }
        ExprKind::Range { start, end, .. } => {
            apply_pascal_ident_renames_expr(start, renames);
            apply_pascal_ident_renames_expr(end, renames);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            apply_pascal_ident_renames_expr(cond, renames);
            apply_pascal_ident_renames_expr(then, renames);
            apply_pascal_ident_renames_expr(else_, renames);
        }
        _ => {}
    }
}

fn default_init_enum_indexed_arrays(
    body: &mut [Statement],
    enum_type_counts: &std::collections::HashMap<String, usize>,
) {
    for stmt in body {
        default_init_enum_indexed_arrays_stmt(stmt, enum_type_counts);
    }
}

fn normalize_pascal_enum_indexed_array_decls(
    body: &mut [Statement],
    enum_type_counts: &std::collections::HashMap<String, usize>,
) {
    for stmt in body {
        match &mut stmt.kind {
            StmtKind::VarDecl { declarations, .. } => {
                for decl in declarations {
                    let Some(type_hint) = decl.type_hint.as_deref() else {
                        continue;
                    };
                    let Some(count) = enum_indexed_array_len(type_hint, enum_type_counts) else {
                        continue;
                    };
                    let Some(element_type) = pascal_array_element_type(type_hint) else {
                        continue;
                    };
                    decl.type_hint = Some(format!("array[0..{}] of {}", count - 1, element_type));
                    decl.array_bounds = Some(vec![
                        Expression::new(ExprKind::Lit(Literal::Int(0))),
                        Expression::new(ExprKind::Lit(Literal::Int(count as i64 - 1))),
                    ]);
                    if decl.init.is_none() {
                        decl.init = Some(null_array_initializer(count));
                    }
                }
            }
            StmtKind::FunctionDecl { body, .. } | StmtKind::Block(body) => {
                normalize_pascal_enum_indexed_array_decls(body, enum_type_counts);
            }
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                normalize_pascal_enum_indexed_array_decls(then_body, enum_type_counts);
                for (_, body) in elifs {
                    normalize_pascal_enum_indexed_array_decls(body, enum_type_counts);
                }
                if let Some(body) = else_body {
                    normalize_pascal_enum_indexed_array_decls(body, enum_type_counts);
                }
            }
            StmtKind::For { body, .. }
            | StmtKind::ForIn { body, .. }
            | StmtKind::While { body, .. }
            | StmtKind::DoWhile { body, .. } => {
                normalize_pascal_enum_indexed_array_decls(body, enum_type_counts);
            }
            _ => {}
        }
    }
}

fn default_init_enum_indexed_arrays_stmt(
    stmt: &mut Statement,
    enum_type_counts: &std::collections::HashMap<String, usize>,
) {
    match &mut stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if decl.init.as_ref().is_some_and(
                    |init| matches!(&init.kind, ExprKind::Array(elements) if elements.is_empty()),
                ) || decl.init.is_none()
                {
                    if let Some(count) = decl
                        .type_hint
                        .as_deref()
                        .and_then(|hint| enum_indexed_array_len(hint, enum_type_counts))
                    {
                        decl.init = Some(null_array_initializer(count));
                        if let Some(type_hint) = decl.type_hint.as_deref() {
                            if let Some(element_type) = pascal_array_element_type(type_hint) {
                                decl.type_hint =
                                    Some(format!("array[0..{}] of {}", count - 1, element_type));
                            }
                        }
                        decl.array_bounds = Some(vec![
                            Expression::new(ExprKind::Lit(Literal::Int(0))),
                            Expression::new(ExprKind::Lit(Literal::Int(count as i64 - 1))),
                        ]);
                    }
                }
            }
        }
        StmtKind::FunctionDecl { body, .. } | StmtKind::Block(body) => {
            for stmt in body {
                default_init_enum_indexed_arrays_stmt(stmt, enum_type_counts);
            }
        }
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            for stmt in then_body {
                default_init_enum_indexed_arrays_stmt(stmt, enum_type_counts);
            }
            for (_, body) in elifs {
                for stmt in body {
                    default_init_enum_indexed_arrays_stmt(stmt, enum_type_counts);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    default_init_enum_indexed_arrays_stmt(stmt, enum_type_counts);
                }
            }
        }
        StmtKind::While { body, .. }
        | StmtKind::DoWhile { body, .. }
        | StmtKind::For { body, .. }
        | StmtKind::ForIn { body, .. } => {
            for stmt in body {
                default_init_enum_indexed_arrays_stmt(stmt, enum_type_counts);
            }
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                match member {
                    ClassMember::Method(method) | ClassMember::NestedType(method) => {
                        default_init_enum_indexed_arrays_stmt(method, enum_type_counts);
                    }
                    ClassMember::Constructor { body, .. } => {
                        for stmt in body {
                            default_init_enum_indexed_arrays_stmt(stmt, enum_type_counts);
                        }
                    }
                    ClassMember::Property { getter, setter, .. } => {
                        if let Some(getter) = getter {
                            for stmt in getter {
                                default_init_enum_indexed_arrays_stmt(stmt, enum_type_counts);
                            }
                        }
                        if let Some(setter) = setter {
                            for stmt in &mut setter.body {
                                default_init_enum_indexed_arrays_stmt(stmt, enum_type_counts);
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

fn enum_indexed_array_len(
    type_hint: &str,
    enum_type_counts: &std::collections::HashMap<String, usize>,
) -> Option<usize> {
    let trimmed = type_hint.trim();
    let lower = trimmed.to_ascii_lowercase();
    let rest = lower.strip_prefix("array[")?;
    let close = rest.find(']')?;
    let index = trimmed["array[".len().."array[".len() + close]
        .trim()
        .to_lowercase();
    enum_type_counts.get(&index).copied()
}

fn rewrite_pascal_enum_ordinals(
    body: &mut [Statement],
    enum_member_ordinals: &std::collections::HashMap<String, i64>,
) {
    for stmt in body {
        rewrite_pascal_enum_ordinals_stmt(stmt, enum_member_ordinals);
    }
}

fn rewrite_pascal_enum_ordinals_stmt(
    stmt: &mut Statement,
    enum_member_ordinals: &std::collections::HashMap<String, i64>,
) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) => rewrite_pascal_enum_ordinals_expr(expr, enum_member_ordinals),
        StmtKind::Return(expr) => {
            if let Some(expr) = expr {
                rewrite_pascal_enum_ordinals_expr(expr, enum_member_ordinals);
            }
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_pascal_enum_ordinals_expr(init, enum_member_ordinals);
                }
            }
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_pascal_enum_ordinals_expr(target, enum_member_ordinals);
            }
            rewrite_pascal_enum_ordinals_expr(value, enum_member_ordinals);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_pascal_enum_ordinals_expr(target, enum_member_ordinals);
            rewrite_pascal_enum_ordinals_expr(value, enum_member_ordinals);
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_pascal_enum_ordinals_expr(cond, enum_member_ordinals);
            for stmt in then_body {
                rewrite_pascal_enum_ordinals_stmt(stmt, enum_member_ordinals);
            }
            for (cond, body) in elifs {
                rewrite_pascal_enum_ordinals_expr(cond, enum_member_ordinals);
                for stmt in body {
                    rewrite_pascal_enum_ordinals_stmt(stmt, enum_member_ordinals);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_pascal_enum_ordinals_stmt(stmt, enum_member_ordinals);
                }
            }
        }
        StmtKind::While { cond, body, .. } | StmtKind::DoWhile { cond, body, .. } => {
            rewrite_pascal_enum_ordinals_expr(cond, enum_member_ordinals);
            for stmt in body {
                rewrite_pascal_enum_ordinals_stmt(stmt, enum_member_ordinals);
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                rewrite_pascal_enum_ordinals_stmt(init, enum_member_ordinals);
            }
            if let Some(cond) = cond {
                rewrite_pascal_enum_ordinals_expr(cond, enum_member_ordinals);
            }
            if let Some(update) = update {
                rewrite_pascal_enum_ordinals_expr(update, enum_member_ordinals);
            }
            for stmt in body {
                rewrite_pascal_enum_ordinals_stmt(stmt, enum_member_ordinals);
            }
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            rewrite_pascal_enum_ordinals_expr(iter, enum_member_ordinals);
            for stmt in body {
                rewrite_pascal_enum_ordinals_stmt(stmt, enum_member_ordinals);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_pascal_enum_ordinals_stmt(stmt, enum_member_ordinals);
                }
            }
        }
        StmtKind::Block(body) | StmtKind::NamespaceDecl { body, .. } => {
            for stmt in body {
                rewrite_pascal_enum_ordinals_stmt(stmt, enum_member_ordinals);
            }
        }
        StmtKind::FunctionDecl { body, .. } => {
            for stmt in body {
                rewrite_pascal_enum_ordinals_stmt(stmt, enum_member_ordinals);
            }
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                match member {
                    ClassMember::Method(method) | ClassMember::NestedType(method) => {
                        rewrite_pascal_enum_ordinals_stmt(method, enum_member_ordinals);
                    }
                    ClassMember::Constructor { body, .. } => {
                        for stmt in body {
                            rewrite_pascal_enum_ordinals_stmt(stmt, enum_member_ordinals);
                        }
                    }
                    ClassMember::Property { getter, setter, .. } => {
                        if let Some(getter) = getter {
                            for stmt in getter {
                                rewrite_pascal_enum_ordinals_stmt(stmt, enum_member_ordinals);
                            }
                        }
                        if let Some(setter) = setter {
                            for stmt in &mut setter.body {
                                rewrite_pascal_enum_ordinals_stmt(stmt, enum_member_ordinals);
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

fn rewrite_pascal_enum_ordinals_expr(
    expr: &mut Expression,
    enum_member_ordinals: &std::collections::HashMap<String, i64>,
) {
    match &mut expr.kind {
        ExprKind::Ident(name) => {
            if let Some(value) = enum_member_ordinals.get(name) {
                *expr = pascal_int(*value);
            }
        }
        ExprKind::Call { callee, args, .. } => {
            rewrite_pascal_enum_ordinals_expr(callee, enum_member_ordinals);
            for arg in args {
                rewrite_pascal_enum_ordinals_expr(&mut arg.value, enum_member_ordinals);
            }
        }
        ExprKind::Member { object, .. } => {
            rewrite_pascal_enum_ordinals_expr(object, enum_member_ordinals)
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_pascal_enum_ordinals_expr(object, enum_member_ordinals);
            rewrite_pascal_enum_ordinals_expr(index, enum_member_ordinals);
        }
        ExprKind::Unary { expr, .. } => {
            rewrite_pascal_enum_ordinals_expr(expr, enum_member_ordinals)
        }
        ExprKind::Binary { left, right, .. } => {
            rewrite_pascal_enum_ordinals_expr(left, enum_member_ordinals);
            rewrite_pascal_enum_ordinals_expr(right, enum_member_ordinals);
        }
        ExprKind::Assign { target, value } | ExprKind::Walrus { target, value } => {
            rewrite_pascal_enum_ordinals_expr(target, enum_member_ordinals);
            rewrite_pascal_enum_ordinals_expr(value, enum_member_ordinals);
        }
        ExprKind::Array(elements) => {
            for element in elements {
                rewrite_pascal_enum_ordinals_expr(&mut element.value, enum_member_ordinals);
                if let Some(key) = &mut element.key {
                    rewrite_pascal_enum_ordinals_expr(key, enum_member_ordinals);
                }
            }
        }
        ExprKind::Set(items) | ExprKind::Tuple(items) | ExprKind::Sequence(items) => {
            for item in items {
                rewrite_pascal_enum_ordinals_expr(item, enum_member_ordinals);
            }
        }
        ExprKind::Object(properties) => {
            for property in properties {
                match property {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        rewrite_pascal_enum_ordinals_expr(key, enum_member_ordinals);
                        rewrite_pascal_enum_ordinals_expr(value, enum_member_ordinals);
                    }
                    ObjectProperty::Spread(value) => {
                        rewrite_pascal_enum_ordinals_expr(value, enum_member_ordinals);
                    }
                    _ => {}
                }
            }
        }
        ExprKind::Range { start, end, .. } => {
            rewrite_pascal_enum_ordinals_expr(start, enum_member_ordinals);
            rewrite_pascal_enum_ordinals_expr(end, enum_member_ordinals);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_pascal_enum_ordinals_expr(cond, enum_member_ordinals);
            rewrite_pascal_enum_ordinals_expr(then, enum_member_ordinals);
            rewrite_pascal_enum_ordinals_expr(else_, enum_member_ordinals);
        }
        _ => {}
    }
}

fn rewrite_pascal_set_semantics(
    body: &mut [Statement],
    enum_type_names: &std::collections::HashSet<String>,
) {
    let set_fields = collect_pascal_set_fields(body);
    let set_param_positions = collect_pascal_set_param_positions(body);
    let mut set_vars = std::collections::HashSet::new();
    for stmt in body {
        rewrite_pascal_set_stmt(
            stmt,
            &mut set_vars,
            &set_fields,
            &set_param_positions,
            enum_type_names,
        );
    }
}

fn rewrite_pascal_set_stmt(
    stmt: &mut Statement,
    set_vars: &mut std::collections::HashSet<String>,
    set_fields: &std::collections::HashSet<String>,
    set_param_positions: &std::collections::HashMap<String, std::collections::HashSet<usize>>,
    enum_type_names: &std::collections::HashSet<String>,
) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) => {
            rewrite_pascal_set_expr(
                expr,
                set_vars,
                set_fields,
                set_param_positions,
                enum_type_names,
            );
            if let Some(rewritten) = pascal_include_exclude_stmt(expr) {
                *stmt = rewritten;
            }
        }
        StmtKind::Return(expr) => {
            if let Some(expr) = expr {
                rewrite_pascal_set_expr(
                    expr,
                    set_vars,
                    set_fields,
                    set_param_positions,
                    enum_type_names,
                );
            }
        }
        StmtKind::Throw { expr, cause } => {
            if let Some(expr) = expr {
                rewrite_pascal_set_expr(
                    expr,
                    set_vars,
                    set_fields,
                    set_param_positions,
                    enum_type_names,
                );
            }
            if let Some(cause) = cause {
                rewrite_pascal_set_expr(
                    cause,
                    set_vars,
                    set_fields,
                    set_param_positions,
                    enum_type_names,
                );
            }
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                let is_set = decl
                    .type_hint
                    .as_deref()
                    .is_some_and(is_pascal_set_type_hint);
                if is_set {
                    if decl.init.is_none() {
                        decl.init = Some(empty_pascal_set_expr());
                    }
                    if let BindingPattern::Ident(name) = &decl.pattern {
                        set_vars.insert(name.to_lowercase());
                    }
                }
                if let Some(init) = &mut decl.init {
                    rewrite_pascal_set_expr(
                        init,
                        set_vars,
                        set_fields,
                        set_param_positions,
                        enum_type_names,
                    );
                }
            }
        }
        StmtKind::Assign { targets, value } => {
            let target_is_set = targets
                .iter()
                .any(|target| is_pascal_set_expr(target, set_vars, set_fields));
            for target in targets {
                rewrite_pascal_set_expr(
                    target,
                    set_vars,
                    set_fields,
                    set_param_positions,
                    enum_type_names,
                );
            }
            rewrite_pascal_set_expr(
                value,
                set_vars,
                set_fields,
                set_param_positions,
                enum_type_names,
            );
            if target_is_set {
                promote_pascal_set_literals_expr(value);
            }
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_pascal_set_expr(
                target,
                set_vars,
                set_fields,
                set_param_positions,
                enum_type_names,
            );
            rewrite_pascal_set_expr(
                value,
                set_vars,
                set_fields,
                set_param_positions,
                enum_type_names,
            );
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_pascal_set_expr(
                cond,
                set_vars,
                set_fields,
                set_param_positions,
                enum_type_names,
            );
            for stmt in then_body {
                rewrite_pascal_set_stmt(
                    stmt,
                    set_vars,
                    set_fields,
                    set_param_positions,
                    enum_type_names,
                );
            }
            for (cond, body) in elifs {
                rewrite_pascal_set_expr(
                    cond,
                    set_vars,
                    set_fields,
                    set_param_positions,
                    enum_type_names,
                );
                for stmt in body {
                    rewrite_pascal_set_stmt(
                        stmt,
                        set_vars,
                        set_fields,
                        set_param_positions,
                        enum_type_names,
                    );
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_pascal_set_stmt(
                        stmt,
                        set_vars,
                        set_fields,
                        set_param_positions,
                        enum_type_names,
                    );
                }
            }
        }
        StmtKind::While { cond, body, .. } | StmtKind::DoWhile { cond, body, .. } => {
            rewrite_pascal_set_expr(
                cond,
                set_vars,
                set_fields,
                set_param_positions,
                enum_type_names,
            );
            for stmt in body {
                rewrite_pascal_set_stmt(
                    stmt,
                    set_vars,
                    set_fields,
                    set_param_positions,
                    enum_type_names,
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
                rewrite_pascal_set_stmt(
                    init,
                    set_vars,
                    set_fields,
                    set_param_positions,
                    enum_type_names,
                );
            }
            if let Some(cond) = cond {
                rewrite_pascal_set_expr(
                    cond,
                    set_vars,
                    set_fields,
                    set_param_positions,
                    enum_type_names,
                );
            }
            if let Some(update) = update {
                rewrite_pascal_set_expr(
                    update,
                    set_vars,
                    set_fields,
                    set_param_positions,
                    enum_type_names,
                );
            }
            for stmt in body {
                rewrite_pascal_set_stmt(
                    stmt,
                    set_vars,
                    set_fields,
                    set_param_positions,
                    enum_type_names,
                );
            }
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            rewrite_pascal_set_expr(
                iter,
                set_vars,
                set_fields,
                set_param_positions,
                enum_type_names,
            );
            for stmt in body {
                rewrite_pascal_set_stmt(
                    stmt,
                    set_vars,
                    set_fields,
                    set_param_positions,
                    enum_type_names,
                );
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_pascal_set_stmt(
                        stmt,
                        set_vars,
                        set_fields,
                        set_param_positions,
                        enum_type_names,
                    );
                }
            }
        }
        StmtKind::Block(body) | StmtKind::NamespaceDecl { body, .. } => {
            for stmt in body {
                rewrite_pascal_set_stmt(
                    stmt,
                    set_vars,
                    set_fields,
                    set_param_positions,
                    enum_type_names,
                );
            }
        }
        StmtKind::FunctionDecl { params, body, .. } => {
            let mut scoped = set_vars.clone();
            for param in params {
                if param
                    .type_hint
                    .as_deref()
                    .is_some_and(is_pascal_set_type_hint)
                {
                    scoped.insert(param.name.to_lowercase());
                }
                if let Some(default) = &mut param.default {
                    rewrite_pascal_set_expr(
                        default,
                        &scoped,
                        set_fields,
                        set_param_positions,
                        enum_type_names,
                    );
                }
            }
            for stmt in body {
                rewrite_pascal_set_stmt(
                    stmt,
                    &mut scoped,
                    set_fields,
                    set_param_positions,
                    enum_type_names,
                );
            }
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                rewrite_pascal_set_member(member, set_fields, set_param_positions, enum_type_names);
            }
        }
        StmtKind::Try {
            body,
            catches,
            finally,
            ..
        } => {
            for stmt in body {
                rewrite_pascal_set_stmt(
                    stmt,
                    set_vars,
                    set_fields,
                    set_param_positions,
                    enum_type_names,
                );
            }
            for catch in catches {
                for stmt in &mut catch.body {
                    rewrite_pascal_set_stmt(
                        stmt,
                        set_vars,
                        set_fields,
                        set_param_positions,
                        enum_type_names,
                    );
                }
            }
            if let Some(body) = finally {
                for stmt in body {
                    rewrite_pascal_set_stmt(
                        stmt,
                        set_vars,
                        set_fields,
                        set_param_positions,
                        enum_type_names,
                    );
                }
            }
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            rewrite_pascal_set_expr(
                expr,
                set_vars,
                set_fields,
                set_param_positions,
                enum_type_names,
            );
            for case in cases {
                for cond in &mut case.conditions {
                    match cond {
                        CaseCondition::Value(expr) | CaseCondition::Comparison { expr, .. } => {
                            rewrite_pascal_set_expr(
                                expr,
                                set_vars,
                                set_fields,
                                set_param_positions,
                                enum_type_names,
                            )
                        }
                        CaseCondition::Range { from, to } => {
                            rewrite_pascal_set_expr(
                                from,
                                set_vars,
                                set_fields,
                                set_param_positions,
                                enum_type_names,
                            );
                            rewrite_pascal_set_expr(
                                to,
                                set_vars,
                                set_fields,
                                set_param_positions,
                                enum_type_names,
                            );
                        }
                    }
                }
                for stmt in &mut case.body {
                    rewrite_pascal_set_stmt(
                        stmt,
                        set_vars,
                        set_fields,
                        set_param_positions,
                        enum_type_names,
                    );
                }
            }
            if let Some(body) = default {
                for stmt in body {
                    rewrite_pascal_set_stmt(
                        stmt,
                        set_vars,
                        set_fields,
                        set_param_positions,
                        enum_type_names,
                    );
                }
            }
        }
        _ => {}
    }
}

fn pascal_include_exclude_stmt(expr: &Expression) -> Option<Statement> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    let ExprKind::Ident(name) = &callee.kind else {
        return None;
    };
    let helper = if name.eq_ignore_ascii_case("Include") {
        "__vybe_pascal_set_include"
    } else if name.eq_ignore_ascii_case("Exclude") {
        "__vybe_pascal_set_exclude"
    } else {
        return None;
    };
    if args.len() != 2 {
        return None;
    }
    let target = args[0].value.clone();
    let value = args[1].value.clone();
    if matches!(target.kind, ExprKind::Member { .. }) {
        let tmp_name = "__pascal_set_member_tmp".to_string();
        return Some(Statement::new(StmtKind::Block(vec![
            Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(tmp_name.clone()),
                    type_hint: None,
                    init: Some(target.clone()),
                    array_bounds: None,
                    with_events: false,
                }],
                kind: VarDeclKind::Dim,
            }),
            Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(helper)),
                args: vec![
                    Argument::positional(Expression::ident(&tmp_name)),
                    Argument::positional(value),
                ],
                optional: false,
            }))),
            Statement::new(StmtKind::Assign {
                targets: vec![target],
                value: Expression::ident(&tmp_name),
            }),
        ])));
    }
    Some(Statement::new(StmtKind::Expr(Expression::new(
        ExprKind::Call {
            callee: Box::new(Expression::ident(helper)),
            args: vec![Argument::positional(target), Argument::positional(value)],
            optional: false,
        },
    ))))
}

fn rewrite_pascal_set_member(
    member: &mut ClassMember,
    set_fields: &std::collections::HashSet<String>,
    set_param_positions: &std::collections::HashMap<String, std::collections::HashSet<usize>>,
    enum_type_names: &std::collections::HashSet<String>,
) {
    let mut scoped = std::collections::HashSet::new();
    match member {
        ClassMember::Method(method) => rewrite_pascal_set_stmt(
            method,
            &mut scoped,
            set_fields,
            set_param_positions,
            enum_type_names,
        ),
        ClassMember::Constructor { params, body, .. } => {
            for param in params {
                if param
                    .type_hint
                    .as_deref()
                    .is_some_and(is_pascal_set_type_hint)
                {
                    scoped.insert(param.name.to_lowercase());
                }
            }
            for stmt in body {
                rewrite_pascal_set_stmt(
                    stmt,
                    &mut scoped,
                    set_fields,
                    set_param_positions,
                    enum_type_names,
                );
            }
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                for stmt in getter {
                    rewrite_pascal_set_stmt(
                        stmt,
                        &mut scoped,
                        set_fields,
                        set_param_positions,
                        enum_type_names,
                    );
                }
            }
            if let Some(setter) = setter {
                for stmt in &mut setter.body {
                    rewrite_pascal_set_stmt(
                        stmt,
                        &mut scoped,
                        set_fields,
                        set_param_positions,
                        enum_type_names,
                    );
                }
            }
        }
        _ => {}
    }
}

fn rewrite_pascal_set_expr(
    expr: &mut Expression,
    set_vars: &std::collections::HashSet<String>,
    set_fields: &std::collections::HashSet<String>,
    set_param_positions: &std::collections::HashMap<String, std::collections::HashSet<usize>>,
    enum_type_names: &std::collections::HashSet<String>,
) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            rewrite_pascal_set_expr(
                callee,
                set_vars,
                set_fields,
                set_param_positions,
                enum_type_names,
            );
            for arg in args.iter_mut() {
                rewrite_pascal_set_expr(
                    &mut arg.value,
                    set_vars,
                    set_fields,
                    set_param_positions,
                    enum_type_names,
                );
            }
            if let ExprKind::Ident(name) = &callee.kind {
                if let Some(positions) = set_param_positions.get(&name.to_lowercase()) {
                    for index in positions {
                        if let Some(arg) = args.get_mut(*index) {
                            promote_pascal_set_literals_expr(&mut arg.value);
                        }
                    }
                }
            }
            if args.len() == 1
                && matches!(&callee.kind, ExprKind::Ident(name) if name.eq_ignore_ascii_case("Length") || name == "__len__")
                && is_pascal_set_expr(&args[0].value, set_vars, set_fields)
            {
                *expr = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__pascal_set_length")),
                    args: vec![Argument::positional(args[0].value.clone())],
                    optional: false,
                });
                return;
            }
            if args.len() == 1
                && matches!(&callee.kind, ExprKind::Ident(name) if enum_type_names.contains(&name.to_lowercase()))
            {
                *expr = args[0].value.clone();
            }
        }
        ExprKind::Member { object, .. } => rewrite_pascal_set_expr(
            object,
            set_vars,
            set_fields,
            set_param_positions,
            enum_type_names,
        ),
        ExprKind::Index { object, index, .. } => {
            rewrite_pascal_set_expr(
                object,
                set_vars,
                set_fields,
                set_param_positions,
                enum_type_names,
            );
            rewrite_pascal_set_expr(
                index,
                set_vars,
                set_fields,
                set_param_positions,
                enum_type_names,
            );
        }
        ExprKind::Unary { expr: inner, .. } => rewrite_pascal_set_expr(
            inner,
            set_vars,
            set_fields,
            set_param_positions,
            enum_type_names,
        ),
        ExprKind::Binary { op, left, right } => {
            rewrite_pascal_set_expr(
                left,
                set_vars,
                set_fields,
                set_param_positions,
                enum_type_names,
            );
            rewrite_pascal_set_expr(
                right,
                set_vars,
                set_fields,
                set_param_positions,
                enum_type_names,
            );
            if *op == BinOp::In && is_pascal_set_expr(right, set_vars, set_fields) {
                *expr = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__vybe_pascal_set_contains")),
                    args: vec![
                        Argument::positional((**left).clone()),
                        Argument::positional((**right).clone()),
                    ],
                    optional: false,
                });
                return;
            }
            if matches!(op, BinOp::Add | BinOp::Sub | BinOp::Mul) {
                if is_pascal_set_expr(left, set_vars, set_fields)
                    || is_pascal_set_expr(right, set_vars, set_fields)
                {
                    promote_pascal_set_literals_expr(left);
                    promote_pascal_set_literals_expr(right);
                }
            }
            if matches!(
                op,
                BinOp::Eq | BinOp::NotEq | BinOp::Lt | BinOp::LtEq | BinOp::Gt | BinOp::GtEq
            ) && (is_pascal_set_expr(left, set_vars, set_fields)
                || is_pascal_set_expr(right, set_vars, set_fields))
            {
                promote_pascal_set_literals_expr(left);
                promote_pascal_set_literals_expr(right);
            }
            if is_pascal_set_expr(left, set_vars, set_fields)
                && is_pascal_set_expr(right, set_vars, set_fields)
            {
                if let Some(rewritten) =
                    pascal_set_comparison_expr(*op, (**left).clone(), (**right).clone())
                {
                    *expr = rewritten;
                }
            }
        }
        ExprKind::Assign { target, value } | ExprKind::Walrus { target, value } => {
            rewrite_pascal_set_expr(
                target,
                set_vars,
                set_fields,
                set_param_positions,
                enum_type_names,
            );
            rewrite_pascal_set_expr(
                value,
                set_vars,
                set_fields,
                set_param_positions,
                enum_type_names,
            );
        }
        ExprKind::Array(elements) => {
            for element in elements {
                rewrite_pascal_set_expr(
                    &mut element.value,
                    set_vars,
                    set_fields,
                    set_param_positions,
                    enum_type_names,
                );
                if let Some(key) = &mut element.key {
                    rewrite_pascal_set_expr(
                        key,
                        set_vars,
                        set_fields,
                        set_param_positions,
                        enum_type_names,
                    );
                }
            }
        }
        ExprKind::Set(items) | ExprKind::Tuple(items) | ExprKind::Sequence(items) => {
            for item in items {
                rewrite_pascal_set_expr(
                    item,
                    set_vars,
                    set_fields,
                    set_param_positions,
                    enum_type_names,
                );
            }
        }
        ExprKind::Object(properties) => {
            for property in properties {
                match property {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        rewrite_pascal_set_expr(
                            key,
                            set_vars,
                            set_fields,
                            set_param_positions,
                            enum_type_names,
                        );
                        rewrite_pascal_set_expr(
                            value,
                            set_vars,
                            set_fields,
                            set_param_positions,
                            enum_type_names,
                        );
                    }
                    ObjectProperty::Spread(value) => {
                        rewrite_pascal_set_expr(
                            value,
                            set_vars,
                            set_fields,
                            set_param_positions,
                            enum_type_names,
                        );
                    }
                    _ => {}
                }
            }
        }
        ExprKind::Range { start, end, .. } => {
            rewrite_pascal_set_expr(
                start,
                set_vars,
                set_fields,
                set_param_positions,
                enum_type_names,
            );
            rewrite_pascal_set_expr(
                end,
                set_vars,
                set_fields,
                set_param_positions,
                enum_type_names,
            );
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_pascal_set_expr(
                cond,
                set_vars,
                set_fields,
                set_param_positions,
                enum_type_names,
            );
            rewrite_pascal_set_expr(
                then,
                set_vars,
                set_fields,
                set_param_positions,
                enum_type_names,
            );
            rewrite_pascal_set_expr(
                else_,
                set_vars,
                set_fields,
                set_param_positions,
                enum_type_names,
            );
        }
        _ => {}
    }
}

fn pascal_set_comparison_expr(
    op: BinOp,
    left: Expression,
    right: Expression,
) -> Option<Expression> {
    match op {
        BinOp::Eq => Some(and_expr(
            pascal_set_subset_expr(left.clone(), right.clone()),
            pascal_set_subset_expr(right, left),
        )),
        BinOp::NotEq => Some(Expression::new(ExprKind::Unary {
            op: UnaryOp::Not,
            expr: Box::new(and_expr(
                pascal_set_subset_expr(left.clone(), right.clone()),
                pascal_set_subset_expr(right, left),
            )),
        })),
        BinOp::LtEq => Some(pascal_set_subset_expr(left, right)),
        BinOp::GtEq => Some(pascal_set_subset_expr(right, left)),
        BinOp::Lt => Some(and_expr(
            pascal_set_subset_expr(left.clone(), right.clone()),
            Expression::new(ExprKind::Unary {
                op: UnaryOp::Not,
                expr: Box::new(pascal_set_subset_expr(right, left)),
            }),
        )),
        BinOp::Gt => Some(and_expr(
            pascal_set_subset_expr(right.clone(), left.clone()),
            Expression::new(ExprKind::Unary {
                op: UnaryOp::Not,
                expr: Box::new(pascal_set_subset_expr(left, right)),
            }),
        )),
        _ => None,
    }
}

fn pascal_set_subset_expr(left: Expression, right: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(high_call(Expression::new(ExprKind::Binary {
            op: BinOp::Sub,
            left: Box::new(left),
            right: Box::new(right),
        }))),
        right: Box::new(pascal_int(-1)),
    })
}

fn and_expr(left: Expression, right: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(left),
        right: Box::new(right),
    })
}

fn high_call(value: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("High")),
        args: vec![Argument::positional(value)],
        optional: false,
    })
}

fn is_pascal_set_expr(
    expr: &Expression,
    set_vars: &std::collections::HashSet<String>,
    set_fields: &std::collections::HashSet<String>,
) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => set_vars.contains(&name.to_lowercase()),
        ExprKind::Member { field, .. } => set_fields.contains(&field.to_lowercase()),
        ExprKind::Set(_) => true,
        ExprKind::Binary { op, left, right } => {
            matches!(op, BinOp::Add | BinOp::Sub | BinOp::Mul)
                && is_pascal_set_expr(left, set_vars, set_fields)
                && is_pascal_set_expr(right, set_vars, set_fields)
        }
        _ => false,
    }
}

fn collect_pascal_set_fields(body: &[Statement]) -> std::collections::HashSet<String> {
    let mut fields = std::collections::HashSet::new();
    for stmt in body {
        match &stmt.kind {
            StmtKind::ClassDecl { members, .. }
            | StmtKind::StructDecl { members, .. }
            | StmtKind::ModuleDecl { members, .. } => {
                for member in members {
                    if let ClassMember::Field {
                        name,
                        type_hint: Some(type_hint),
                        ..
                    } = member
                    {
                        if is_pascal_set_type_hint(type_hint) {
                            fields.insert(name.to_lowercase());
                        }
                    }
                }
            }
            StmtKind::Block(nested) | StmtKind::NamespaceDecl { body: nested, .. } => {
                fields.extend(collect_pascal_set_fields(nested));
            }
            _ => {}
        }
    }
    fields
}

fn collect_pascal_set_param_positions(
    body: &[Statement],
) -> std::collections::HashMap<String, std::collections::HashSet<usize>> {
    let mut out = std::collections::HashMap::new();
    for stmt in body {
        match &stmt.kind {
            StmtKind::FunctionDecl { name, params, .. } => {
                let positions = params
                    .iter()
                    .enumerate()
                    .filter_map(|(index, param)| {
                        param
                            .type_hint
                            .as_deref()
                            .is_some_and(is_pascal_set_type_hint)
                            .then_some(index)
                    })
                    .collect::<std::collections::HashSet<_>>();
                if !positions.is_empty() {
                    out.insert(name.to_lowercase(), positions);
                }
            }
            StmtKind::Block(nested) | StmtKind::NamespaceDecl { body: nested, .. } => {
                out.extend(collect_pascal_set_param_positions(nested));
            }
            _ => {}
        }
    }
    out
}

fn promote_pascal_set_literals_expr(expr: &mut Expression) {
    match &mut expr.kind {
        ExprKind::Array(elements) => {
            if elements
                .iter()
                .all(|element| !element.spread && element.key.is_none())
            {
                let items = elements
                    .iter_mut()
                    .map(|element| {
                        promote_pascal_set_literals_expr(&mut element.value);
                        element.value.clone()
                    })
                    .collect();
                expr.kind = ExprKind::Set(items);
            } else {
                for element in elements {
                    promote_pascal_set_literals_expr(&mut element.value);
                    if let Some(key) = &mut element.key {
                        promote_pascal_set_literals_expr(key);
                    }
                }
            }
        }
        ExprKind::Binary { left, right, .. } => {
            promote_pascal_set_literals_expr(left);
            promote_pascal_set_literals_expr(right);
        }
        ExprKind::Unary { expr, .. } => promote_pascal_set_literals_expr(expr),
        ExprKind::Call { callee, args, .. } => {
            promote_pascal_set_literals_expr(callee);
            for arg in args {
                promote_pascal_set_literals_expr(&mut arg.value);
            }
        }
        ExprKind::Set(items) | ExprKind::Tuple(items) | ExprKind::Sequence(items) => {
            for item in items {
                promote_pascal_set_literals_expr(item);
            }
        }
        ExprKind::Member { object, .. } => promote_pascal_set_literals_expr(object),
        ExprKind::Index { object, index, .. } => {
            promote_pascal_set_literals_expr(object);
            promote_pascal_set_literals_expr(index);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            promote_pascal_set_literals_expr(cond);
            promote_pascal_set_literals_expr(then);
            promote_pascal_set_literals_expr(else_);
        }
        _ => {}
    }
}

fn is_pascal_set_type_hint(type_hint: &str) -> bool {
    normalize_pascal_type_hint(type_hint)
        .to_ascii_lowercase()
        .starts_with("set of ")
}

fn empty_pascal_set_expr() -> Expression {
    Expression::new(ExprKind::Set(Vec::new()))
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
                    name: None,
                    params: vec![msg_param],
                    body: vec![assign_msg],
                    base_args: None,
                    initializer_target: vybe_ast::ConstructorInitializerTarget::Base,
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

fn collect_static_members(
    body: &[Statement],
) -> (
    std::collections::HashSet<(String, String)>,
    std::collections::HashSet<(String, String)>,
) {
    let mut methods_out = std::collections::HashSet::new();
    let mut values_out = std::collections::HashSet::new();
    for stmt in body {
        if let StmtKind::ClassDecl { name, members, .. } = &stmt.kind {
            for member in members {
                match member {
                    ClassMember::Method(method) => {
                        let StmtKind::FunctionDecl {
                            name: method_name,
                            modifiers,
                            ..
                        } = &method.kind
                        else {
                            continue;
                        };
                        if modifiers.is_static {
                            methods_out.insert((name.to_lowercase(), method_name.to_lowercase()));
                        }
                    }
                    ClassMember::Const {
                        name: member_name, ..
                    } => {
                        values_out.insert((name.to_lowercase(), member_name.to_lowercase()));
                    }
                    ClassMember::Property {
                        name: member_name,
                        modifiers,
                        ..
                    } if modifiers.is_static => {
                        values_out.insert((name.to_lowercase(), member_name.to_lowercase()));
                    }
                    _ => {}
                }
            }
        }
    }
    (methods_out, values_out)
}

fn collect_static_var_param_indices(
    body: &[Statement],
) -> std::collections::HashMap<(String, String), std::collections::HashSet<usize>> {
    let mut out = std::collections::HashMap::new();
    for stmt in body {
        let StmtKind::ClassDecl { name, members, .. } = &stmt.kind else {
            continue;
        };
        for member in members {
            let ClassMember::Method(method) = member else {
                continue;
            };
            let StmtKind::FunctionDecl {
                name: method_name,
                params,
                modifiers,
                ..
            } = &method.kind
            else {
                continue;
            };
            if !modifiers.is_static {
                continue;
            }
            let indexes: std::collections::HashSet<usize> = params
                .iter()
                .enumerate()
                .filter_map(|(idx, param)| {
                    matches!(param.pass_by, PassBy::Ref | PassBy::Out).then_some(idx)
                })
                .collect();
            if !indexes.is_empty() {
                out.insert((name.to_lowercase(), method_name.to_lowercase()), indexes);
            }
        }
    }
    out
}

fn mark_static_var_args_stmt(
    stmt: &mut Statement,
    static_var_params: &std::collections::HashMap<
        (String, String),
        std::collections::HashSet<usize>,
    >,
) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) => mark_static_var_args_expr(expr, static_var_params),
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    mark_static_var_args_expr(init, static_var_params);
                }
            }
        }
        StmtKind::FunctionDecl { body, .. } | StmtKind::Block(body) => {
            for stmt in body {
                mark_static_var_args_stmt(stmt, static_var_params);
            }
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                mark_static_var_args_member(member, static_var_params);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            mark_static_var_args_expr(cond, static_var_params);
            for stmt in then_body {
                mark_static_var_args_stmt(stmt, static_var_params);
            }
            for (cond, body) in elifs {
                mark_static_var_args_expr(cond, static_var_params);
                for stmt in body {
                    mark_static_var_args_stmt(stmt, static_var_params);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    mark_static_var_args_stmt(stmt, static_var_params);
                }
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            mark_static_var_args_expr(cond, static_var_params);
            for stmt in body {
                mark_static_var_args_stmt(stmt, static_var_params);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    mark_static_var_args_stmt(stmt, static_var_params);
                }
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            for stmt in body {
                mark_static_var_args_stmt(stmt, static_var_params);
            }
            mark_static_var_args_expr(cond, static_var_params);
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                mark_static_var_args_stmt(init, static_var_params);
            }
            if let Some(cond) = cond {
                mark_static_var_args_expr(cond, static_var_params);
            }
            if let Some(update) = update {
                mark_static_var_args_expr(update, static_var_params);
            }
            for stmt in body {
                mark_static_var_args_stmt(stmt, static_var_params);
            }
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            mark_static_var_args_expr(iter, static_var_params);
            for stmt in body {
                mark_static_var_args_stmt(stmt, static_var_params);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    mark_static_var_args_stmt(stmt, static_var_params);
                }
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            for stmt in body {
                mark_static_var_args_stmt(stmt, static_var_params);
            }
            for catch in catches {
                for stmt in &mut catch.body {
                    mark_static_var_args_stmt(stmt, static_var_params);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    mark_static_var_args_stmt(stmt, static_var_params);
                }
            }
            if let Some(body) = finally {
                for stmt in body {
                    mark_static_var_args_stmt(stmt, static_var_params);
                }
            }
        }
        StmtKind::With { items, body, .. } => {
            for item in items {
                mark_static_var_args_expr(&mut item.expr, static_var_params);
            }
            for stmt in body {
                mark_static_var_args_stmt(stmt, static_var_params);
            }
        }
        StmtKind::Return(Some(expr))
        | StmtKind::Throw {
            expr: Some(expr), ..
        } => {
            mark_static_var_args_expr(expr, static_var_params);
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                mark_static_var_args_expr(target, static_var_params);
            }
            mark_static_var_args_expr(value, static_var_params);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            mark_static_var_args_expr(target, static_var_params);
            mark_static_var_args_expr(value, static_var_params);
        }
        _ => {}
    }
}

fn mark_static_var_args_member(
    member: &mut ClassMember,
    static_var_params: &std::collections::HashMap<
        (String, String),
        std::collections::HashSet<usize>,
    >,
) {
    match member {
        ClassMember::Field {
            init: Some(expr), ..
        }
        | ClassMember::Const { value: expr, .. } => {
            mark_static_var_args_expr(expr, static_var_params)
        }
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            mark_static_var_args_stmt(stmt, static_var_params);
        }
        ClassMember::Constructor { body, .. } => {
            for stmt in body {
                mark_static_var_args_stmt(stmt, static_var_params);
            }
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                for stmt in getter {
                    mark_static_var_args_stmt(stmt, static_var_params);
                }
            }
            if let Some(setter) = setter {
                for stmt in &mut setter.body {
                    mark_static_var_args_stmt(stmt, static_var_params);
                }
            }
        }
        _ => {}
    }
}

fn mark_static_var_args_expr(
    expr: &mut Expression,
    static_var_params: &std::collections::HashMap<
        (String, String),
        std::collections::HashSet<usize>,
    >,
) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                if let ExprKind::Ident(class_name) = &object.kind {
                    if let Some(indexes) =
                        static_var_params.get(&(class_name.to_lowercase(), field.to_lowercase()))
                    {
                        for (idx, arg) in args.iter_mut().enumerate() {
                            if indexes.contains(&idx) {
                                arg.by_ref = true;
                            }
                        }
                    }
                }
            }
            for arg in args {
                mark_static_var_args_expr(&mut arg.value, static_var_params);
            }
        }
        ExprKind::Binary { left, right, .. } => {
            mark_static_var_args_expr(left, static_var_params);
            mark_static_var_args_expr(right, static_var_params);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Await(expr)
        | ExprKind::Yield(Some(expr))
        | ExprKind::Spread(expr)
        | ExprKind::TypeOf(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Cast { expr, .. } => mark_static_var_args_expr(expr, static_var_params),
        ExprKind::Ternary { cond, then, else_ } => {
            mark_static_var_args_expr(cond, static_var_params);
            mark_static_var_args_expr(then, static_var_params);
            mark_static_var_args_expr(else_, static_var_params);
        }
        ExprKind::Member { object, .. } => mark_static_var_args_expr(object, static_var_params),
        ExprKind::Index { object, index, .. } => {
            mark_static_var_args_expr(object, static_var_params);
            mark_static_var_args_expr(index, static_var_params);
        }
        ExprKind::New { class, args } => {
            mark_static_var_args_expr(class, static_var_params);
            for arg in args {
                mark_static_var_args_expr(&mut arg.value, static_var_params);
            }
        }
        ExprKind::Array(elements) => {
            for element in elements {
                if let Some(key) = &mut element.key {
                    mark_static_var_args_expr(key, static_var_params);
                }
                mark_static_var_args_expr(&mut element.value, static_var_params);
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) => {
            for item in items {
                mark_static_var_args_expr(item, static_var_params);
            }
        }
        ExprKind::NullCoalesce { left, right } => {
            mark_static_var_args_expr(left, static_var_params);
            mark_static_var_args_expr(right, static_var_params);
        }
        ExprKind::Range { start, end, .. } => {
            mark_static_var_args_expr(start, static_var_params);
            mark_static_var_args_expr(end, static_var_params);
        }
        ExprKind::Assign { target, value } | ExprKind::Walrus { target, value } => {
            mark_static_var_args_expr(target, static_var_params);
            mark_static_var_args_expr(value, static_var_params);
        }
        ExprKind::SuperCall { args, .. } => {
            for arg in args {
                mark_static_var_args_expr(&mut arg.value, static_var_params);
            }
        }
        ExprKind::StaticAccess { class, member } => {
            mark_static_var_args_expr(class, static_var_params);
            mark_static_var_args_expr(member, static_var_params);
        }
        _ => {}
    }
}

fn pascal_type_stamp_expr(object: Expression) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(object),
        field: "__type".to_string(),
        null_safe: false,
    })
}

fn pascal_class_name_expr(object: Expression, class_names: &[(String, String)]) -> Expression {
    let type_expr = pascal_type_stamp_expr(object);
    class_names
        .iter()
        .rev()
        .fold(type_expr.clone(), |fallback, (canonical, display)| {
            Expression::new(ExprKind::Ternary {
                cond: Box::new(Expression::new(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(type_expr.clone()),
                    right: Box::new(Expression::string(canonical)),
                })),
                then: Box::new(Expression::string(display)),
                else_: Box::new(fallback),
            })
        })
}

fn pascal_class_ref_expr(object: Expression, class_names: &[(String, String)]) -> Expression {
    let type_expr = pascal_type_stamp_expr(object);
    class_names
        .iter()
        .rev()
        .fold(Expression::null(), |fallback, (canonical, display)| {
            Expression::new(ExprKind::Ternary {
                cond: Box::new(Expression::new(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(type_expr.clone()),
                    right: Box::new(Expression::string(canonical)),
                })),
                then: Box::new(Expression::ident(display)),
                else_: Box::new(fallback),
            })
        })
}

fn rewrite_pascal_rtti_stmt(stmt: &mut Statement, class_names: &[(String, String)]) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) => rewrite_pascal_rtti_expr(expr, class_names),
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_pascal_rtti_expr(init, class_names);
                }
            }
        }
        StmtKind::FunctionDecl { body, .. } | StmtKind::Block(body) => {
            for stmt in body {
                rewrite_pascal_rtti_stmt(stmt, class_names);
            }
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                rewrite_pascal_rtti_member(member, class_names);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_pascal_rtti_expr(cond, class_names);
            for stmt in then_body {
                rewrite_pascal_rtti_stmt(stmt, class_names);
            }
            for (cond, body) in elifs {
                rewrite_pascal_rtti_expr(cond, class_names);
                for stmt in body {
                    rewrite_pascal_rtti_stmt(stmt, class_names);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_pascal_rtti_stmt(stmt, class_names);
                }
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            rewrite_pascal_rtti_expr(cond, class_names);
            for stmt in body {
                rewrite_pascal_rtti_stmt(stmt, class_names);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_pascal_rtti_stmt(stmt, class_names);
                }
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            for stmt in body {
                rewrite_pascal_rtti_stmt(stmt, class_names);
            }
            rewrite_pascal_rtti_expr(cond, class_names);
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                rewrite_pascal_rtti_stmt(init, class_names);
            }
            if let Some(cond) = cond {
                rewrite_pascal_rtti_expr(cond, class_names);
            }
            if let Some(update) = update {
                rewrite_pascal_rtti_expr(update, class_names);
            }
            for stmt in body {
                rewrite_pascal_rtti_stmt(stmt, class_names);
            }
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            rewrite_pascal_rtti_expr(iter, class_names);
            for stmt in body {
                rewrite_pascal_rtti_stmt(stmt, class_names);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_pascal_rtti_stmt(stmt, class_names);
                }
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            for stmt in body {
                rewrite_pascal_rtti_stmt(stmt, class_names);
            }
            for catch in catches {
                for stmt in &mut catch.body {
                    rewrite_pascal_rtti_stmt(stmt, class_names);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_pascal_rtti_stmt(stmt, class_names);
                }
            }
            if let Some(body) = finally {
                for stmt in body {
                    rewrite_pascal_rtti_stmt(stmt, class_names);
                }
            }
        }
        StmtKind::With { items, body, .. } => {
            for item in items {
                rewrite_pascal_rtti_expr(&mut item.expr, class_names);
            }
            for stmt in body {
                rewrite_pascal_rtti_stmt(stmt, class_names);
            }
        }
        StmtKind::Return(Some(expr))
        | StmtKind::Throw {
            expr: Some(expr), ..
        } => {
            rewrite_pascal_rtti_expr(expr, class_names);
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_pascal_rtti_expr(target, class_names);
            }
            rewrite_pascal_rtti_expr(value, class_names);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_pascal_rtti_expr(target, class_names);
            rewrite_pascal_rtti_expr(value, class_names);
        }
        _ => {}
    }
}

fn rewrite_pascal_rtti_member(member: &mut ClassMember, class_names: &[(String, String)]) {
    match member {
        ClassMember::Field {
            init: Some(expr), ..
        }
        | ClassMember::Const { value: expr, .. } => rewrite_pascal_rtti_expr(expr, class_names),
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            rewrite_pascal_rtti_stmt(stmt, class_names);
        }
        ClassMember::Constructor { body, .. } => {
            for stmt in body {
                rewrite_pascal_rtti_stmt(stmt, class_names);
            }
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                for stmt in getter {
                    rewrite_pascal_rtti_stmt(stmt, class_names);
                }
            }
            if let Some(setter) = setter {
                for stmt in &mut setter.body {
                    rewrite_pascal_rtti_stmt(stmt, class_names);
                }
            }
        }
        _ => {}
    }
}

fn rewrite_pascal_rtti_expr(expr: &mut Expression, class_names: &[(String, String)]) {
    match &mut expr.kind {
        ExprKind::Member { object, field, .. }
            if field == "__pascal_class_name" || field == "__pascal_class_type" =>
        {
            rewrite_pascal_rtti_expr(object, class_names);
            *expr = if field == "__pascal_class_name" {
                pascal_class_name_expr((**object).clone(), class_names)
            } else {
                pascal_class_ref_expr((**object).clone(), class_names)
            };
        }
        ExprKind::Call { callee, args, .. } => {
            rewrite_pascal_rtti_expr(callee, class_names);
            for arg in args {
                rewrite_pascal_rtti_expr(&mut arg.value, class_names);
            }
        }
        ExprKind::Binary { left, right, .. } => {
            rewrite_pascal_rtti_expr(left, class_names);
            rewrite_pascal_rtti_expr(right, class_names);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Await(expr)
        | ExprKind::Yield(Some(expr))
        | ExprKind::Spread(expr)
        | ExprKind::TypeOf(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Cast { expr, .. } => rewrite_pascal_rtti_expr(expr, class_names),
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_pascal_rtti_expr(cond, class_names);
            rewrite_pascal_rtti_expr(then, class_names);
            rewrite_pascal_rtti_expr(else_, class_names);
        }
        ExprKind::Member { object, .. } => rewrite_pascal_rtti_expr(object, class_names),
        ExprKind::Index { object, index, .. } => {
            rewrite_pascal_rtti_expr(object, class_names);
            rewrite_pascal_rtti_expr(index, class_names);
        }
        ExprKind::New { class, args } => {
            rewrite_pascal_rtti_expr(class, class_names);
            for arg in args {
                rewrite_pascal_rtti_expr(&mut arg.value, class_names);
            }
        }
        ExprKind::Array(elements) => {
            for element in elements {
                if let Some(key) = &mut element.key {
                    rewrite_pascal_rtti_expr(key, class_names);
                }
                rewrite_pascal_rtti_expr(&mut element.value, class_names);
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) => {
            for item in items {
                rewrite_pascal_rtti_expr(item, class_names);
            }
        }
        ExprKind::NullCoalesce { left, right } => {
            rewrite_pascal_rtti_expr(left, class_names);
            rewrite_pascal_rtti_expr(right, class_names);
        }
        ExprKind::Range { start, end, .. } => {
            rewrite_pascal_rtti_expr(start, class_names);
            rewrite_pascal_rtti_expr(end, class_names);
        }
        ExprKind::Assign { target, value } | ExprKind::Walrus { target, value } => {
            rewrite_pascal_rtti_expr(target, class_names);
            rewrite_pascal_rtti_expr(value, class_names);
        }
        ExprKind::SuperCall { args, .. } => {
            for arg in args {
                rewrite_pascal_rtti_expr(&mut arg.value, class_names);
            }
        }
        ExprKind::StaticAccess { class, member } => {
            rewrite_pascal_rtti_expr(class, class_names);
            rewrite_pascal_rtti_expr(member, class_names);
        }
        _ => {}
    }
}

fn collect_zero_arg_instance_methods(body: &[Statement]) -> std::collections::HashSet<String> {
    let mut out = std::collections::HashSet::new();
    for stmt in body {
        if let StmtKind::ClassDecl { members, .. } | StmtKind::StructDecl { members, .. } =
            &stmt.kind
        {
            for member in members {
                let ClassMember::Method(method) = member else {
                    continue;
                };
                let StmtKind::FunctionDecl {
                    name,
                    params,
                    modifiers,
                    ..
                } = &method.kind
                else {
                    continue;
                };
                if params.is_empty() && !modifiers.is_static {
                    out.insert(name.to_lowercase());
                }
            }
        }
    }
    out
}

fn rewrite_zero_arg_instance_method_refs_stmt(
    stmt: &mut Statement,
    methods: &std::collections::HashSet<String>,
) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) => rewrite_zero_arg_instance_method_refs_expr(expr, methods),
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_zero_arg_instance_method_refs_expr(init, methods);
                }
            }
        }
        StmtKind::FunctionDecl { params, body, .. } => {
            let scoped = method_names_without_params(methods, params);
            for stmt in body {
                rewrite_zero_arg_instance_method_refs_stmt(stmt, &scoped);
            }
        }
        StmtKind::Block(body) => {
            for stmt in body {
                rewrite_zero_arg_instance_method_refs_stmt(stmt, methods);
            }
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                rewrite_zero_arg_instance_method_refs_member(member, methods);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_zero_arg_instance_method_refs_expr(cond, methods);
            for stmt in then_body {
                rewrite_zero_arg_instance_method_refs_stmt(stmt, methods);
            }
            for (cond, body) in elifs {
                rewrite_zero_arg_instance_method_refs_expr(cond, methods);
                for stmt in body {
                    rewrite_zero_arg_instance_method_refs_stmt(stmt, methods);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_zero_arg_instance_method_refs_stmt(stmt, methods);
                }
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            rewrite_zero_arg_instance_method_refs_expr(cond, methods);
            for stmt in body {
                rewrite_zero_arg_instance_method_refs_stmt(stmt, methods);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_zero_arg_instance_method_refs_stmt(stmt, methods);
                }
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            for stmt in body {
                rewrite_zero_arg_instance_method_refs_stmt(stmt, methods);
            }
            rewrite_zero_arg_instance_method_refs_expr(cond, methods);
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                rewrite_zero_arg_instance_method_refs_stmt(init, methods);
            }
            if let Some(cond) = cond {
                rewrite_zero_arg_instance_method_refs_expr(cond, methods);
            }
            if let Some(update) = update {
                rewrite_zero_arg_instance_method_refs_expr(update, methods);
            }
            for stmt in body {
                rewrite_zero_arg_instance_method_refs_stmt(stmt, methods);
            }
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            rewrite_zero_arg_instance_method_refs_expr(iter, methods);
            for stmt in body {
                rewrite_zero_arg_instance_method_refs_stmt(stmt, methods);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_zero_arg_instance_method_refs_stmt(stmt, methods);
                }
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            for stmt in body {
                rewrite_zero_arg_instance_method_refs_stmt(stmt, methods);
            }
            for catch in catches {
                for stmt in &mut catch.body {
                    rewrite_zero_arg_instance_method_refs_stmt(stmt, methods);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_zero_arg_instance_method_refs_stmt(stmt, methods);
                }
            }
            if let Some(body) = finally {
                for stmt in body {
                    rewrite_zero_arg_instance_method_refs_stmt(stmt, methods);
                }
            }
        }
        StmtKind::With { items, body, .. } => {
            for item in items {
                rewrite_zero_arg_instance_method_refs_expr(&mut item.expr, methods);
            }
            for stmt in body {
                rewrite_zero_arg_instance_method_refs_stmt(stmt, methods);
            }
        }
        StmtKind::Return(Some(expr))
        | StmtKind::Throw {
            expr: Some(expr), ..
        } => {
            rewrite_zero_arg_instance_method_refs_expr(expr, methods);
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_zero_arg_instance_method_refs_expr(target, methods);
            }
            rewrite_zero_arg_instance_method_refs_expr(value, methods);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_zero_arg_instance_method_refs_expr(target, methods);
            rewrite_zero_arg_instance_method_refs_expr(value, methods);
        }
        _ => {}
    }
}

fn rewrite_zero_arg_instance_method_refs_member(
    member: &mut ClassMember,
    methods: &std::collections::HashSet<String>,
) {
    match member {
        ClassMember::Field {
            init: Some(expr), ..
        }
        | ClassMember::Const { value: expr, .. } => {
            rewrite_zero_arg_instance_method_refs_expr(expr, methods)
        }
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            rewrite_zero_arg_instance_method_refs_stmt(stmt, methods);
        }
        ClassMember::Constructor { params, body, .. } => {
            let scoped = method_names_without_params(methods, params);
            for stmt in body {
                rewrite_zero_arg_instance_method_refs_stmt(stmt, &scoped);
            }
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                for stmt in getter {
                    rewrite_zero_arg_instance_method_refs_stmt(stmt, methods);
                }
            }
            if let Some(setter) = setter {
                for stmt in &mut setter.body {
                    rewrite_zero_arg_instance_method_refs_stmt(stmt, methods);
                }
            }
        }
        _ => {}
    }
}

fn rewrite_zero_arg_instance_method_refs_expr(
    expr: &mut Expression,
    methods: &std::collections::HashSet<String>,
) {
    match &mut expr.kind {
        ExprKind::Call { args, .. } => {
            for arg in args {
                rewrite_zero_arg_instance_method_refs_expr(&mut arg.value, methods);
            }
        }
        ExprKind::Member { object, field, .. } => {
            rewrite_zero_arg_instance_method_refs_expr(object, methods);
            if methods.contains(&field.to_lowercase()) {
                expr.kind = ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Member {
                        object: Box::new((**object).clone()),
                        field: field.clone(),
                        null_safe: false,
                    })),
                    args: Vec::new(),
                    optional: false,
                };
            }
        }
        ExprKind::Binary { left, right, .. } => {
            rewrite_zero_arg_instance_method_refs_expr(left, methods);
            rewrite_zero_arg_instance_method_refs_expr(right, methods);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Await(expr)
        | ExprKind::Yield(Some(expr))
        | ExprKind::Spread(expr)
        | ExprKind::TypeOf(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Cast { expr, .. } => rewrite_zero_arg_instance_method_refs_expr(expr, methods),
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_zero_arg_instance_method_refs_expr(cond, methods);
            rewrite_zero_arg_instance_method_refs_expr(then, methods);
            rewrite_zero_arg_instance_method_refs_expr(else_, methods);
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_zero_arg_instance_method_refs_expr(object, methods);
            rewrite_zero_arg_instance_method_refs_expr(index, methods);
        }
        ExprKind::New { class, args } => {
            rewrite_zero_arg_instance_method_refs_expr(class, methods);
            for arg in args {
                rewrite_zero_arg_instance_method_refs_expr(&mut arg.value, methods);
            }
        }
        ExprKind::Array(elements) => {
            for element in elements {
                if let Some(key) = &mut element.key {
                    rewrite_zero_arg_instance_method_refs_expr(key, methods);
                }
                rewrite_zero_arg_instance_method_refs_expr(&mut element.value, methods);
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) => {
            for item in items {
                rewrite_zero_arg_instance_method_refs_expr(item, methods);
            }
        }
        ExprKind::NullCoalesce { left, right } => {
            rewrite_zero_arg_instance_method_refs_expr(left, methods);
            rewrite_zero_arg_instance_method_refs_expr(right, methods);
        }
        ExprKind::Range { start, end, .. } => {
            rewrite_zero_arg_instance_method_refs_expr(start, methods);
            rewrite_zero_arg_instance_method_refs_expr(end, methods);
        }
        ExprKind::Assign { target, value } | ExprKind::Walrus { target, value } => {
            rewrite_zero_arg_instance_method_refs_expr(target, methods);
            rewrite_zero_arg_instance_method_refs_expr(value, methods);
        }
        ExprKind::SuperCall { args, .. } => {
            for arg in args {
                rewrite_zero_arg_instance_method_refs_expr(&mut arg.value, methods);
            }
        }
        ExprKind::StaticAccess { class, member } => {
            rewrite_zero_arg_instance_method_refs_expr(class, methods);
            rewrite_zero_arg_instance_method_refs_expr(member, methods);
        }
        _ => {}
    }
}

fn rewrite_static_method_calls_stmt(
    stmt: &mut Statement,
    methods: &std::collections::HashSet<(String, String)>,
    values: &std::collections::HashSet<(String, String)>,
) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) => rewrite_static_method_calls_expr(expr, methods, values),
        StmtKind::Block(body) | StmtKind::NamespaceDecl { body, .. } => {
            for stmt in body {
                rewrite_static_method_calls_stmt(stmt, methods, values);
            }
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_static_method_calls_expr(init, methods, values);
                }
            }
        }
        StmtKind::FunctionDecl { body, .. } => {
            for stmt in body {
                rewrite_static_method_calls_stmt(stmt, methods, values);
            }
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                rewrite_static_method_calls_member(member, methods, values);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_static_method_calls_expr(cond, methods, values);
            for stmt in then_body {
                rewrite_static_method_calls_stmt(stmt, methods, values);
            }
            for (cond, body) in elifs {
                rewrite_static_method_calls_expr(cond, methods, values);
                for stmt in body {
                    rewrite_static_method_calls_stmt(stmt, methods, values);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_static_method_calls_stmt(stmt, methods, values);
                }
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            rewrite_static_method_calls_expr(cond, methods, values);
            for stmt in body {
                rewrite_static_method_calls_stmt(stmt, methods, values);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_static_method_calls_stmt(stmt, methods, values);
                }
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            for stmt in body {
                rewrite_static_method_calls_stmt(stmt, methods, values);
            }
            rewrite_static_method_calls_expr(cond, methods, values);
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                rewrite_static_method_calls_stmt(init, methods, values);
            }
            if let Some(cond) = cond {
                rewrite_static_method_calls_expr(cond, methods, values);
            }
            if let Some(update) = update {
                rewrite_static_method_calls_expr(update, methods, values);
            }
            for stmt in body {
                rewrite_static_method_calls_stmt(stmt, methods, values);
            }
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            rewrite_static_method_calls_expr(iter, methods, values);
            for stmt in body {
                rewrite_static_method_calls_stmt(stmt, methods, values);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_static_method_calls_stmt(stmt, methods, values);
                }
            }
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            rewrite_static_method_calls_expr(expr, methods, values);
            for case in cases {
                for stmt in &mut case.body {
                    rewrite_static_method_calls_stmt(stmt, methods, values);
                }
            }
            if let Some(default) = default {
                for stmt in default {
                    rewrite_static_method_calls_stmt(stmt, methods, values);
                }
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            for stmt in body {
                rewrite_static_method_calls_stmt(stmt, methods, values);
            }
            for catch in catches {
                for stmt in &mut catch.body {
                    rewrite_static_method_calls_stmt(stmt, methods, values);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_static_method_calls_stmt(stmt, methods, values);
                }
            }
            if let Some(body) = finally {
                for stmt in body {
                    rewrite_static_method_calls_stmt(stmt, methods, values);
                }
            }
        }
        StmtKind::With { items, body, .. } => {
            for item in items {
                rewrite_static_method_calls_expr(&mut item.expr, methods, values);
            }
            for stmt in body {
                rewrite_static_method_calls_stmt(stmt, methods, values);
            }
        }
        StmtKind::Return(Some(expr))
        | StmtKind::Throw {
            expr: Some(expr), ..
        } => {
            rewrite_static_method_calls_expr(expr, methods, values);
        }
        StmtKind::Throw {
            cause: Some(cause), ..
        } => rewrite_static_method_calls_expr(cause, methods, values),
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_static_method_calls_expr(target, methods, values);
            }
            rewrite_static_method_calls_expr(value, methods, values);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_static_method_calls_expr(target, methods, values);
            rewrite_static_method_calls_expr(value, methods, values);
        }
        _ => {}
    }
}

fn rewrite_static_method_calls_member(
    member: &mut ClassMember,
    methods: &std::collections::HashSet<(String, String)>,
    values: &std::collections::HashSet<(String, String)>,
) {
    match member {
        ClassMember::Field {
            init: Some(expr), ..
        }
        | ClassMember::Const { value: expr, .. } => {
            rewrite_static_method_calls_expr(expr, methods, values)
        }
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            rewrite_static_method_calls_stmt(stmt, methods, values)
        }
        ClassMember::Constructor { body, .. } => {
            for stmt in body {
                rewrite_static_method_calls_stmt(stmt, methods, values);
            }
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                for stmt in getter {
                    rewrite_static_method_calls_stmt(stmt, methods, values);
                }
            }
            if let Some(setter) = setter {
                for stmt in &mut setter.body {
                    rewrite_static_method_calls_stmt(stmt, methods, values);
                }
            }
        }
        _ => {}
    }
}

fn rewrite_static_method_calls_expr(
    expr: &mut Expression,
    methods: &std::collections::HashSet<(String, String)>,
    values: &std::collections::HashSet<(String, String)>,
) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            for arg in args.iter_mut() {
                rewrite_static_method_calls_expr(&mut arg.value, methods, values);
            }
            if !matches!(
                &callee.kind,
                ExprKind::Member { object, field, .. }
                    if matches!(&object.kind, ExprKind::Ident(class_name)
                        if methods.contains(&(class_name.to_lowercase(), field.to_lowercase())))
            ) {
                rewrite_static_method_calls_expr(callee, methods, values);
            }
        }
        ExprKind::Binary { left, right, .. } => {
            rewrite_static_method_calls_expr(left, methods, values);
            rewrite_static_method_calls_expr(right, methods, values);
        }
        ExprKind::Unary { expr: inner, .. }
        | ExprKind::Await(inner)
        | ExprKind::Yield(Some(inner))
        | ExprKind::Spread(inner) => rewrite_static_method_calls_expr(inner, methods, values),
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_static_method_calls_expr(cond, methods, values);
            rewrite_static_method_calls_expr(then, methods, values);
            rewrite_static_method_calls_expr(else_, methods, values);
        }
        ExprKind::Member { object, field, .. } => {
            rewrite_static_method_calls_expr(object, methods, values);
            if let ExprKind::Ident(class_name) = &object.kind {
                if methods.contains(&(class_name.to_lowercase(), field.to_lowercase())) {
                    expr.kind = ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Member {
                            object: Box::new(Expression::ident(class_name)),
                            field: field.clone(),
                            null_safe: false,
                        })),
                        args: Vec::new(),
                        optional: false,
                    };
                } else if values.contains(&(class_name.to_lowercase(), field.to_lowercase())) {
                    expr.kind = ExprKind::StaticAccess {
                        class: Box::new(Expression::ident(class_name)),
                        member: Box::new(Expression::ident(field)),
                    };
                }
            }
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_static_method_calls_expr(object, methods, values);
            rewrite_static_method_calls_expr(index, methods, values);
        }
        ExprKind::New { args, .. } => {
            for arg in args {
                rewrite_static_method_calls_expr(&mut arg.value, methods, values);
            }
        }
        ExprKind::Array(elements) => {
            for element in elements {
                if let Some(key) = &mut element.key {
                    rewrite_static_method_calls_expr(key, methods, values);
                }
                rewrite_static_method_calls_expr(&mut element.value, methods, values);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value, .. } => {
                        rewrite_static_method_calls_expr(key, methods, values);
                        rewrite_static_method_calls_expr(value, methods, values);
                    }
                    ObjectProperty::Spread(value) => {
                        rewrite_static_method_calls_expr(value, methods, values);
                    }
                    _ => {}
                }
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) => {
            for item in items {
                rewrite_static_method_calls_expr(item, methods, values);
            }
        }
        ExprKind::Cast { expr: inner, .. }
        | ExprKind::TypeOf(inner)
        | ExprKind::Void(inner)
        | ExprKind::Delete(inner)
        | ExprKind::YieldFrom(inner) => rewrite_static_method_calls_expr(inner, methods, values),
        ExprKind::NullCoalesce { left, right } => {
            rewrite_static_method_calls_expr(left, methods, values);
            rewrite_static_method_calls_expr(right, methods, values);
        }
        ExprKind::Range { start, end, .. } => {
            rewrite_static_method_calls_expr(start, methods, values);
            rewrite_static_method_calls_expr(end, methods, values);
        }
        ExprKind::Assign { target, value } | ExprKind::Walrus { target, value } => {
            rewrite_static_method_calls_expr(target, methods, values);
            rewrite_static_method_calls_expr(value, methods, values);
        }
        ExprKind::SuperCall { args, .. } => {
            for arg in args {
                rewrite_static_method_calls_expr(&mut arg.value, methods, values);
            }
        }
        _ => {}
    }
}

/// Walk a statement and rewrite `ClassName.Create(args)` into `New { class, args }`
/// when `ClassName` matches a class declared in the same module.
fn rewrite_constructor_calls_stmt(
    stmt: &mut Statement,
    classes: &std::collections::HashSet<String>,
    static_methods: &std::collections::HashSet<(String, String)>,
) {
    match &mut stmt.kind {
        StmtKind::Expr(e) => rewrite_constructor_calls_expr(e, classes, static_methods),
        StmtKind::Block(stmts) => {
            for s in stmts {
                rewrite_constructor_calls_stmt(s, classes, static_methods);
            }
        }
        StmtKind::VarDecl { declarations, .. } => {
            for d in declarations {
                if let Some(e) = &mut d.init {
                    rewrite_constructor_calls_expr(e, classes, static_methods);
                }
            }
        }
        StmtKind::FunctionDecl { body, .. } => {
            for s in body {
                rewrite_constructor_calls_stmt(s, classes, static_methods);
            }
        }
        StmtKind::ClassDecl { members, .. } => {
            for m in members {
                rewrite_constructor_calls_member(m, classes, static_methods);
            }
        }
        StmtKind::StructDecl { members, .. } | StmtKind::ModuleDecl { members, .. } => {
            for m in members {
                rewrite_constructor_calls_member(m, classes, static_methods);
            }
        }
        StmtKind::NamespaceDecl { body, .. } => {
            for s in body {
                rewrite_constructor_calls_stmt(s, classes, static_methods);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_constructor_calls_expr(cond, classes, static_methods);
            for s in then_body {
                rewrite_constructor_calls_stmt(s, classes, static_methods);
            }
            for (c, b) in elifs {
                rewrite_constructor_calls_expr(c, classes, static_methods);
                for s in b {
                    rewrite_constructor_calls_stmt(s, classes, static_methods);
                }
            }
            if let Some(b) = else_body {
                for s in b {
                    rewrite_constructor_calls_stmt(s, classes, static_methods);
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
                rewrite_constructor_calls_stmt(i, classes, static_methods);
            }
            if let Some(c) = cond {
                rewrite_constructor_calls_expr(c, classes, static_methods);
            }
            if let Some(u) = update {
                rewrite_constructor_calls_expr(u, classes, static_methods);
            }
            for s in body {
                rewrite_constructor_calls_stmt(s, classes, static_methods);
            }
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            rewrite_constructor_calls_expr(iter, classes, static_methods);
            for s in body {
                rewrite_constructor_calls_stmt(s, classes, static_methods);
            }
            if let Some(b) = else_body {
                for s in b {
                    rewrite_constructor_calls_stmt(s, classes, static_methods);
                }
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            rewrite_constructor_calls_expr(cond, classes, static_methods);
            for s in body {
                rewrite_constructor_calls_stmt(s, classes, static_methods);
            }
            if let Some(b) = else_body {
                for s in b {
                    rewrite_constructor_calls_stmt(s, classes, static_methods);
                }
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            for s in body {
                rewrite_constructor_calls_stmt(s, classes, static_methods);
            }
            rewrite_constructor_calls_expr(cond, classes, static_methods);
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            rewrite_constructor_calls_expr(expr, classes, static_methods);
            for c in cases {
                for s in &mut c.body {
                    rewrite_constructor_calls_stmt(s, classes, static_methods);
                }
            }
            if let Some(b) = default {
                for s in b {
                    rewrite_constructor_calls_stmt(s, classes, static_methods);
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
                rewrite_constructor_calls_stmt(s, classes, static_methods);
            }
            for c in catches {
                for s in &mut c.body {
                    rewrite_constructor_calls_stmt(s, classes, static_methods);
                }
            }
            if let Some(b) = else_body {
                for s in b {
                    rewrite_constructor_calls_stmt(s, classes, static_methods);
                }
            }
            if let Some(b) = finally {
                for s in b {
                    rewrite_constructor_calls_stmt(s, classes, static_methods);
                }
            }
        }
        StmtKind::With { items, body, .. } => {
            for it in items {
                rewrite_constructor_calls_expr(&mut it.expr, classes, static_methods);
            }
            for s in body {
                rewrite_constructor_calls_stmt(s, classes, static_methods);
            }
        }
        StmtKind::Return(Some(e)) => rewrite_constructor_calls_expr(e, classes, static_methods),
        StmtKind::Throw { expr: Some(e), .. } => {
            rewrite_constructor_calls_expr(e, classes, static_methods)
        }
        StmtKind::Assign { targets, value } => {
            for t in targets {
                rewrite_constructor_calls_expr(t, classes, static_methods);
            }
            rewrite_constructor_calls_expr(value, classes, static_methods);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_constructor_calls_expr(target, classes, static_methods);
            rewrite_constructor_calls_expr(value, classes, static_methods);
        }
        _ => {}
    }
}

fn rewrite_constructor_calls_member(
    m: &mut ClassMember,
    classes: &std::collections::HashSet<String>,
    static_methods: &std::collections::HashSet<(String, String)>,
) {
    match m {
        ClassMember::Field { init: Some(e), .. } => {
            rewrite_constructor_calls_expr(e, classes, static_methods)
        }
        ClassMember::Method(stmt) => rewrite_constructor_calls_stmt(stmt, classes, static_methods),
        ClassMember::Constructor { body, .. } => {
            for s in body {
                rewrite_constructor_calls_stmt(s, classes, static_methods);
            }
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(g) = getter {
                for s in g {
                    rewrite_constructor_calls_stmt(s, classes, static_methods);
                }
            }
            if let Some(set) = setter {
                for s in &mut set.body {
                    rewrite_constructor_calls_stmt(s, classes, static_methods);
                }
            }
        }
        ClassMember::Const { value, .. } => {
            rewrite_constructor_calls_expr(value, classes, static_methods)
        }
        ClassMember::NestedType(stmt) => {
            rewrite_constructor_calls_stmt(stmt, classes, static_methods)
        }
        _ => {}
    }
}

fn rewrite_constructor_calls_expr(
    expr: &mut Expression,
    classes: &std::collections::HashSet<String>,
    static_methods: &std::collections::HashSet<(String, String)>,
) {
    // Check Call(Member(ClassName, "Create"), args) BEFORE descending so the
    // Member-only rewrite below doesn't fire on the callee position first and
    // turn `TFoo.Create(42)` into a call on a New expression.
    if let ExprKind::Call { callee, args, .. } = &expr.kind {
        if let ExprKind::Ident(class_name) = &callee.kind {
            if classes.contains(&class_name.to_lowercase()) && args.len() == 1 {
                let mut value = args[0].value.clone();
                rewrite_constructor_calls_expr(&mut value, classes, static_methods);
                *expr = value;
                return;
            }
        }
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            if let ExprKind::Ident(class_name) = &object.kind {
                let is_declared_static_method =
                    static_methods.contains(&(class_name.to_lowercase(), field.to_lowercase()));
                let is_ctor_name = field
                    .get(..6)
                    .map_or(false, |prefix| prefix.eq_ignore_ascii_case("Create"));
                if classes.contains(&class_name.to_lowercase())
                    && is_ctor_name
                    && !is_declared_static_method
                {
                    let new_class = Box::new(Expression::ident(class_name));
                    let mut new_args = args.clone();
                    for a in new_args.iter_mut() {
                        rewrite_constructor_calls_expr(&mut a.value, classes, static_methods);
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
            rewrite_constructor_calls_expr(left, classes, static_methods);
            rewrite_constructor_calls_expr(right, classes, static_methods);
        }
        ExprKind::Unary { expr: e, .. } => {
            rewrite_constructor_calls_expr(e, classes, static_methods)
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_constructor_calls_expr(cond, classes, static_methods);
            rewrite_constructor_calls_expr(then, classes, static_methods);
            rewrite_constructor_calls_expr(else_, classes, static_methods);
        }
        ExprKind::Member { object, .. } => {
            rewrite_constructor_calls_expr(object, classes, static_methods)
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_constructor_calls_expr(object, classes, static_methods);
            rewrite_constructor_calls_expr(index, classes, static_methods);
        }
        ExprKind::Call { callee, args, .. } => {
            rewrite_constructor_calls_expr(callee, classes, static_methods);
            for a in args.iter_mut() {
                rewrite_constructor_calls_expr(&mut a.value, classes, static_methods);
            }
        }
        ExprKind::New { class, args } => {
            rewrite_constructor_calls_expr(class, classes, static_methods);
            for a in args.iter_mut() {
                rewrite_constructor_calls_expr(&mut a.value, classes, static_methods);
            }
        }
        ExprKind::Assign { target, value } => {
            rewrite_constructor_calls_expr(target, classes, static_methods);
            rewrite_constructor_calls_expr(value, classes, static_methods);
        }
        ExprKind::Array(elems) => {
            for el in elems {
                rewrite_constructor_calls_expr(&mut el.value, classes, static_methods);
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => {
            for e in items {
                rewrite_constructor_calls_expr(e, classes, static_methods);
            }
        }
        ExprKind::Object(props) => {
            for p in props {
                if let ObjectProperty::KeyValue { value, .. } = p {
                    rewrite_constructor_calls_expr(value, classes, static_methods);
                }
            }
        }
        ExprKind::Interpolation(parts) => {
            for p in parts {
                match p {
                    InterpolPart::Expr(e) | InterpolPart::Formatted(e, _) => {
                        rewrite_constructor_calls_expr(e, classes, static_methods)
                    }
                    _ => {}
                }
            }
        }
        ExprKind::NullCoalesce { left, right } => {
            rewrite_constructor_calls_expr(left, classes, static_methods);
            rewrite_constructor_calls_expr(right, classes, static_methods);
        }
        ExprKind::Range { start, end, .. } => {
            rewrite_constructor_calls_expr(start, classes, static_methods);
            rewrite_constructor_calls_expr(end, classes, static_methods);
        }
        ExprKind::IsType { expr: e, .. } | ExprKind::Cast { expr: e, .. } => {
            rewrite_constructor_calls_expr(e, classes, static_methods)
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

    if let ExprKind::Cast {
        expr: inner,
        type_name,
    } = &expr.kind
    {
        if classes.contains(&type_name.to_lowercase()) {
            *expr = (**inner).clone();
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

        // Try constructor first: Pascal constructors are commonly named Create,
        // but named constructors such as CreateWithCode are valid too.
        let is_create = method_name
            .get(..6)
            .map_or(false, |prefix| prefix.eq_ignore_ascii_case("Create"));
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
                            let mut merged_modifiers = mods.clone();
                            merged_modifiers.is_static |= mm.is_static;
                            merged_modifiers.is_virtual |= mm.is_virtual;
                            merged_modifiers.is_override |= mm.is_override;
                            merged_modifiers.is_abstract |= mm.is_abstract;
                            merged_modifiers.is_overloads |= mm.is_overloads;
                            merged_modifiers.visibility = mm.visibility.clone();
                            *mm = merged_modifiers;
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

fn normalize_pascal_gcl_form_classes(body: &mut [Statement]) {
    for stmt in body {
        let StmtKind::ClassDecl {
            name,
            parents,
            members,
            ..
        } = &mut stmt.kind
        else {
            continue;
        };
        if !parents
            .iter()
            .any(|parent| parent.eq_ignore_ascii_case("TForm"))
        {
            continue;
        }

        let default_name = default_form_instance_name(name);
        let name_stmt = gcl_form_name_assignment(&default_name);
        let mut has_constructor = false;
        for member in members.iter_mut() {
            let ClassMember::Constructor {
                params,
                body,
                base_args,
                ..
            } = member
            else {
                continue;
            };
            has_constructor = true;
            if params.is_empty() {
                params.push(gcl_owner_param());
            }
            if base_args.is_none() {
                *base_args = Some(vec![Expression::ident("AOwner")]);
            }
            if !body.iter().any(is_gcl_form_name_assignment) {
                body.insert(0, name_stmt.clone());
            }
            for stmt in body.iter_mut() {
                normalize_gcl_form_property_stmt(stmt);
            }
        }

        if !has_constructor {
            members.push(ClassMember::Constructor {
                name: None,
                params: vec![gcl_owner_param()],
                body: vec![name_stmt],
                base_args: Some(vec![Expression::ident("AOwner")]),
                initializer_target: ConstructorInitializerTarget::Base,
                visibility: Visibility::Public,
            });
        }
    }
}

fn normalize_pascal_gcl_exprs(body: &mut [Statement]) {
    for stmt in body {
        normalize_gcl_form_property_stmt(stmt);
        match &mut stmt.kind {
            StmtKind::ClassDecl { members, .. }
            | StmtKind::StructDecl { members, .. }
            | StmtKind::ModuleDecl { members, .. } => {
                for member in members {
                    normalize_pascal_gcl_exprs_member(member);
                }
            }
            _ => {}
        }
    }
}

fn normalize_pascal_gcl_exprs_member(member: &mut ClassMember) {
    match member {
        ClassMember::Method(stmt) => normalize_gcl_form_property_stmt(stmt),
        ClassMember::Constructor { body, .. } => normalize_gcl_form_property_stmts(body),
        ClassMember::Property {
            getter,
            setter: Some(setter),
            ..
        } => {
            if let Some(getter) = getter {
                normalize_gcl_form_property_stmts(getter);
            }
            normalize_gcl_form_property_stmts(&mut setter.body);
        }
        ClassMember::Property {
            getter: Some(getter),
            setter: None,
            ..
        } => normalize_gcl_form_property_stmts(getter),
        _ => {}
    }
}

fn normalize_gcl_form_property_stmt(stmt: &mut Statement) {
    if let Some(rewritten) = normalize_gcl_create_form_stmt(stmt) {
        *stmt = rewritten;
        return;
    }
    match &mut stmt.kind {
        StmtKind::Expr(expr) => normalize_gcl_form_property_expr(expr),
        StmtKind::Assign { targets, value } => {
            for target in targets {
                normalize_gcl_form_property_target(target);
            }
            normalize_gcl_form_property_expr(value);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            normalize_gcl_form_property_target(target);
            normalize_gcl_form_property_expr(value);
        }
        StmtKind::Block(stmts) => {
            for stmt in stmts {
                normalize_gcl_form_property_stmt(stmt);
            }
        }
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            normalize_gcl_form_property_stmts(then_body);
            for (_, body) in elifs {
                normalize_gcl_form_property_stmts(body);
            }
            if let Some(body) = else_body {
                normalize_gcl_form_property_stmts(body);
            }
        }
        StmtKind::While {
            body, else_body, ..
        } => {
            normalize_gcl_form_property_stmts(body);
            if let Some(body) = else_body {
                normalize_gcl_form_property_stmts(body);
            }
        }
        StmtKind::DoWhile { body, .. }
        | StmtKind::For { body, .. }
        | StmtKind::ForIn { body, .. } => normalize_gcl_form_property_stmts(body),
        _ => {}
    }
}

fn normalize_gcl_create_form_stmt(stmt: &Statement) -> Option<Statement> {
    let StmtKind::Expr(expr) = &stmt.kind else {
        return None;
    };
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if args.len() != 2 {
        return None;
    }
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    if !field.eq_ignore_ascii_case("CreateForm") {
        return None;
    }
    let ExprKind::Ident(app_name) = &object.kind else {
        return None;
    };
    if !app_name.eq_ignore_ascii_case("Application") {
        return None;
    }
    let ExprKind::Ident(target_name) = &args[1].value.kind else {
        return None;
    };
    let assign_form = Statement::with_span(
        StmtKind::Assign {
            targets: vec![Expression::ident(target_name)],
            value: Expression::new(ExprKind::New {
                class: Box::new(args[0].value.clone()),
                args: vec![Argument::positional(Expression::new(ExprKind::Lit(
                    Literal::Null,
                )))],
            }),
        },
        stmt.span.clone(),
    );
    let remember_main_form = Statement::with_span(
        StmtKind::Assign {
            targets: vec![Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident("Application")),
                field: "__main_form".to_string(),
                null_safe: false,
            })],
            value: Expression::ident(target_name),
        },
        stmt.span.clone(),
    );
    Some(Statement::with_span(
        StmtKind::Block(vec![assign_form, remember_main_form]),
        stmt.span.clone(),
    ))
}

fn normalize_gcl_form_property_target(target: &mut Expression) {
    if let ExprKind::Ident(name) = &target.kind {
        if is_gcl_form_property(name) {
            *target = Expression::new(ExprKind::Member {
                object: Box::new(Expression::new(ExprKind::This)),
                field: name.clone(),
                null_safe: false,
            });
        }
    }
}

fn normalize_gcl_form_property_stmts(stmts: &mut [Statement]) {
    for stmt in stmts {
        normalize_gcl_form_property_stmt(stmt);
    }
}

fn normalize_gcl_form_property_expr(expr: &mut Expression) {
    if let ExprKind::Call { callee, args, .. } = &mut expr.kind {
        if matches!(&callee.kind, ExprKind::Ident(name) if name.eq_ignore_ascii_case("TextToShortCut"))
            && args.len() == 1
        {
            *expr = args[0].value.clone();
            return;
        }
        normalize_gcl_dotted_menu_add_call(callee);
    }

    match &mut expr.kind {
        ExprKind::Assign { target, value } => {
            if let ExprKind::Ident(name) = &target.kind {
                if is_gcl_form_property(name) {
                    *target = Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(Expression::new(ExprKind::This)),
                        field: name.clone(),
                        null_safe: false,
                    }));
                }
            }
            normalize_gcl_form_property_expr(value);
        }
        ExprKind::Call { callee, args, .. } => {
            normalize_gcl_menu_add_call(callee);
            normalize_gcl_form_property_expr(callee);
            for arg in args {
                normalize_gcl_form_property_expr(&mut arg.value);
            }
        }
        ExprKind::Member { object, .. } => normalize_gcl_form_property_expr(object),
        ExprKind::Binary { left, right, .. } => {
            normalize_gcl_form_property_expr(left);
            normalize_gcl_form_property_expr(right);
        }
        ExprKind::Unary { expr, .. } => normalize_gcl_form_property_expr(expr),
        _ => {}
    }
}

fn normalize_gcl_menu_add_call(callee: &mut Box<Expression>) {
    let ExprKind::Member { object, field, .. } = &mut callee.kind else {
        return;
    };
    if !field.eq_ignore_ascii_case("Add") {
        return;
    }
    if matches!(
        &object.kind,
        ExprKind::Member { field, .. }
            if field.eq_ignore_ascii_case("Items")
                || field.eq_ignore_ascii_case("Controls")
                || field.eq_ignore_ascii_case("Components")
    ) {
        return;
    }
    let is_menu_like = match &object.kind {
        ExprKind::Ident(name) => is_menu_like_name(name),
        ExprKind::Member { field, .. } => is_menu_like_name(field),
        _ => false,
    };
    if is_menu_like {
        let original = (**object).clone();
        *object = Box::new(Expression::new(ExprKind::Member {
            object: Box::new(original),
            field: "Items".to_string(),
            null_safe: false,
        }));
    }
}

fn normalize_gcl_dotted_menu_add_call(callee: &mut Box<Expression>) {
    let ExprKind::Ident(name) = &callee.kind else {
        return;
    };
    let Some((receiver, method)) = name.rsplit_once('.') else {
        return;
    };
    if !method.eq_ignore_ascii_case("Add") || !is_menu_like_name(receiver) {
        return;
    }
    *callee = Box::new(Expression::new(ExprKind::Member {
        object: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident(receiver)),
            field: "Items".to_string(),
            null_safe: false,
        })),
        field: "Add".to_string(),
        null_safe: false,
    }));
}

fn is_menu_like_name(name: &str) -> bool {
    let lower = name.to_ascii_lowercase();
    lower.contains("menu") || lower.ends_with("item")
}

fn is_gcl_form_property(name: &str) -> bool {
    matches!(
        name.to_ascii_lowercase().as_str(),
        "caption" | "clientwidth" | "clientheight" | "oncreate" | "onclose"
    )
}

fn default_form_instance_name(class_name: &str) -> String {
    let stripped = class_name
        .strip_prefix('T')
        .filter(|rest| rest.chars().next().is_some_and(|c| c.is_ascii_uppercase()))
        .unwrap_or(class_name);
    stripped.to_lowercase()
}

fn gcl_owner_param() -> Param {
    Param {
        name: "AOwner".to_string(),
        type_hint: Some("TObject".to_string()),
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: true,
    }
}

fn gcl_form_name_assignment(default_name: &str) -> Statement {
    Statement::new(StmtKind::Expr(Expression::new(ExprKind::Assign {
        target: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(Expression::new(ExprKind::This)),
            field: "Name".to_string(),
            null_safe: false,
        })),
        value: Box::new(Expression::new(ExprKind::Lit(Literal::Str(
            default_name.to_string(),
        )))),
    })))
}

fn is_gcl_form_name_assignment(stmt: &Statement) -> bool {
    matches!(
        &stmt.kind,
        StmtKind::Expr(Expression {
            kind: ExprKind::Assign { target, .. },
            ..
        }) if matches!(
            &target.kind,
            ExprKind::Member { field, .. } if field.eq_ignore_ascii_case("Name")
        )
    )
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

fn lower_pascal_gotos_in_body(body: &mut Vec<Statement>) {
    for stmt in body.iter_mut() {
        lower_pascal_gotos_in_stmt(stmt);
    }
    *body = lower_pascal_goto_block(std::mem::take(body));
}

fn lower_pascal_gotos_in_stmt(stmt: &mut Statement) {
    match &mut stmt.kind {
        StmtKind::Block(stmts) => lower_pascal_gotos_in_body(stmts),
        StmtKind::FunctionDecl { body, .. } => lower_pascal_gotos_in_body(body),
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                lower_pascal_gotos_in_member(member);
            }
        }
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            lower_pascal_gotos_in_body(then_body);
            for (_, body) in elifs {
                lower_pascal_gotos_in_body(body);
            }
            if let Some(body) = else_body {
                lower_pascal_gotos_in_body(body);
            }
        }
        StmtKind::For { init, body, .. } => {
            if let Some(init) = init {
                lower_pascal_gotos_in_stmt(init);
            }
            lower_pascal_gotos_in_body(body);
        }
        StmtKind::ForIn {
            body, else_body, ..
        }
        | StmtKind::While {
            body, else_body, ..
        } => {
            lower_pascal_gotos_in_body(body);
            if let Some(body) = else_body {
                lower_pascal_gotos_in_body(body);
            }
        }
        StmtKind::DoWhile { body, .. } => lower_pascal_gotos_in_body(body),
        StmtKind::Switch { cases, default, .. } => {
            for case in cases {
                lower_pascal_gotos_in_body(&mut case.body);
            }
            if let Some(body) = default {
                lower_pascal_gotos_in_body(body);
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            lower_pascal_gotos_in_body(body);
            for catch in catches {
                lower_pascal_gotos_in_body(&mut catch.body);
            }
            if let Some(body) = else_body {
                lower_pascal_gotos_in_body(body);
            }
            if let Some(body) = finally {
                lower_pascal_gotos_in_body(body);
            }
        }
        StmtKind::With { body, .. } => lower_pascal_gotos_in_body(body),
        _ => {}
    }
}

fn lower_pascal_gotos_in_member(member: &mut ClassMember) {
    match member {
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            lower_pascal_gotos_in_stmt(stmt)
        }
        ClassMember::Constructor { body, .. } => lower_pascal_gotos_in_body(body),
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                lower_pascal_gotos_in_body(getter);
            }
            if let Some(setter) = setter {
                lower_pascal_gotos_in_body(&mut setter.body);
            }
        }
        _ => {}
    }
}

fn lower_pascal_goto_block(body: Vec<Statement>) -> Vec<Statement> {
    let mut label_to_block = std::collections::HashMap::new();
    let mut blocks: Vec<Vec<Statement>> = vec![Vec::new()];

    for stmt in body {
        if let StmtKind::Label(name) = stmt.kind {
            let idx = blocks.len() as i64;
            label_to_block.insert(name.to_lowercase(), idx);
            blocks.push(Vec::new());
        } else if let Some(block) = blocks.last_mut() {
            block.push(stmt);
        }
    }

    if label_to_block.is_empty() {
        return blocks.into_iter().next().unwrap_or_default();
    }

    let mut prelude = Vec::new();
    if let Some(first) = blocks.first_mut() {
        while first
            .first()
            .map(is_pascal_goto_prelude_stmt)
            .unwrap_or(false)
        {
            prelude.push(first.remove(0));
        }
    }

    let dispatch_label = "__pascal_goto_dispatch".to_string();
    let pc_name = "__pascal_goto_pc".to_string();
    let total_blocks = blocks.len();
    let mut cases = Vec::new();

    for (idx, block) in blocks.into_iter().enumerate() {
        let next_pc = if idx + 1 < total_blocks {
            pascal_int((idx + 1) as i64)
        } else {
            pascal_int(-1)
        };
        let mut case_body = vec![pascal_assign_stmt(&pc_name, next_pc)];
        case_body.extend(rewrite_pascal_gotos_in_stmts(
            block,
            &label_to_block,
            &pc_name,
            &dispatch_label,
        ));
        case_body.push(Statement::new(StmtKind::Break(BreakTarget::Implicit)));
        cases.push(SwitchCase {
            conditions: vec![CaseCondition::Value(pascal_int(idx as i64))],
            body: case_body,
        });
    }

    let while_body = vec![
        Statement::new(StmtKind::Switch {
            expr: Expression::ident(&pc_name),
            cases,
            default: Some(vec![Statement::new(StmtKind::Break(BreakTarget::Implicit))]),
        }),
        Statement::new(StmtKind::If {
            cond: Expression::new(ExprKind::Binary {
                op: BinOp::Lt,
                left: Box::new(Expression::ident(&pc_name)),
                right: Box::new(pascal_int(0)),
            }),
            then_body: vec![Statement::new(StmtKind::Break(BreakTarget::Implicit))],
            elifs: Vec::new(),
            else_body: None,
        }),
    ];

    prelude.push(Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(pc_name.clone()),
            type_hint: Some("Integer".to_string()),
            init: Some(pascal_int(0)),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Dim,
    }));
    prelude.push(Statement::new(StmtKind::Labeled {
        label: dispatch_label,
        body: Box::new(Statement::new(StmtKind::While {
            cond: Expression::new(ExprKind::Lit(Literal::Bool(true))),
            body: while_body,
            else_body: None,
        })),
    }));
    prelude
}

fn is_pascal_goto_prelude_stmt(stmt: &Statement) -> bool {
    matches!(
        stmt.kind,
        StmtKind::VarDecl { .. }
            | StmtKind::FunctionDecl { .. }
            | StmtKind::ClassDecl { .. }
            | StmtKind::StructDecl { .. }
            | StmtKind::ModuleDecl { .. }
    )
}

fn rewrite_pascal_gotos_in_stmts(
    stmts: Vec<Statement>,
    label_to_block: &std::collections::HashMap<String, i64>,
    pc_name: &str,
    dispatch_label: &str,
) -> Vec<Statement> {
    stmts
        .into_iter()
        .flat_map(|stmt| {
            rewrite_pascal_gotos_in_stmt(stmt, label_to_block, pc_name, dispatch_label)
        })
        .collect()
}

fn rewrite_pascal_gotos_in_stmt(
    stmt: Statement,
    label_to_block: &std::collections::HashMap<String, i64>,
    pc_name: &str,
    dispatch_label: &str,
) -> Vec<Statement> {
    match stmt.kind {
        StmtKind::GoTo(target) => label_to_block
            .get(&target.to_lowercase())
            .map(|idx| {
                vec![
                    pascal_assign_stmt(pc_name, pascal_int(*idx)),
                    Statement::new(StmtKind::Continue(ContinueTarget::Label(
                        dispatch_label.to_string(),
                    ))),
                ]
            })
            .unwrap_or_default(),
        StmtKind::Block(body) => vec![Statement::new(StmtKind::Block(
            rewrite_pascal_gotos_in_stmts(body, label_to_block, pc_name, dispatch_label),
        ))],
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => vec![Statement::new(StmtKind::If {
            cond,
            then_body: rewrite_pascal_gotos_in_stmts(
                then_body,
                label_to_block,
                pc_name,
                dispatch_label,
            ),
            elifs: elifs
                .into_iter()
                .map(|(cond, body)| {
                    (
                        cond,
                        rewrite_pascal_gotos_in_stmts(
                            body,
                            label_to_block,
                            pc_name,
                            dispatch_label,
                        ),
                    )
                })
                .collect(),
            else_body: else_body.map(|body| {
                rewrite_pascal_gotos_in_stmts(body, label_to_block, pc_name, dispatch_label)
            }),
        })],
        StmtKind::While {
            cond,
            body,
            else_body,
        } => vec![Statement::new(StmtKind::While {
            cond,
            body: rewrite_pascal_gotos_in_stmts(body, label_to_block, pc_name, dispatch_label),
            else_body: else_body.map(|body| {
                rewrite_pascal_gotos_in_stmts(body, label_to_block, pc_name, dispatch_label)
            }),
        })],
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => vec![Statement::new(StmtKind::For {
            init,
            cond,
            update,
            body: rewrite_pascal_gotos_in_stmts(body, label_to_block, pc_name, dispatch_label),
        })],
        other => vec![Statement::new(other)],
    }
}

fn pascal_int(value: i64) -> Expression {
    Expression::new(ExprKind::Lit(Literal::Int(value)))
}

fn pascal_assign_stmt(name: &str, value: Expression) -> Statement {
    Statement::new(StmtKind::Assign {
        targets: vec![Expression::ident(name)],
        value,
    })
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
            Rule::label_section => {}
            Rule::class_var_decl_impl => {
                body.push(walk_class_var_decl_impl(decl)?);
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
                let _ = extract_array_bounds(&p)?;
                type_hint = Some(type_ref_to_string(&p));
            }
            Rule::var_init | Rule::inline_var_init => {
                for inner in p.into_inner() {
                    if inner.as_rule() == Rule::expression {
                        init = Some(walk_expression(inner)?);
                    }
                }
            }
            _ => {}
        }
    }

    Ok(build_var_declarators(names, type_hint, init, None))
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
    let mut modifiers = ClassModifiers::default();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::class_modifier => match p.as_str().to_ascii_lowercase().as_str() {
                "abstract" => modifiers.is_abstract = true,
                "sealed" => modifiers.is_sealed = true,
                "static" => modifiers.is_static = true,
                _ => {}
            },
            Rule::class_heritage => {
                for ty in p.into_inner() {
                    if ty.as_rule() == Rule::type_ref {
                        parents.push(type_ref_to_string(&ty));
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
            modifiers,
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
            Rule::class_const_section => members.extend(walk_class_const_section(m)?),
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

fn walk_class_const_section(pair: Pair<Rule>) -> Result<Vec<ClassMember>, String> {
    let mut members = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::class_const_decl {
            members.push(walk_class_const_decl(p)?);
        }
    }
    Ok(members)
}

fn walk_class_const_decl(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut name = String::new();
    let mut type_hint: Option<String> = None;
    let mut value: Option<Expression> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => name = p.as_str().to_string(),
            Rule::type_ref => type_hint = Some(type_ref_to_string(&p)),
            Rule::expression => value = Some(walk_expression(p)?),
            _ => {}
        }
    }

    Ok(ClassMember::Const {
        name,
        type_hint,
        value: value.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Undefined))),
        visibility: Visibility::Public,
    })
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
                    Rule::method_name => name = pascal_method_name_text(sp),
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
    let mut array_bounds: Option<Vec<Expression>> = None;

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
            _ => {}
        }
    }
    let init = type_hint.as_deref().and_then(default_field_init_for_type);

    Ok(names
        .into_iter()
        .map(|n| ClassMember::Field {
            name: n,
            type_hint: type_hint.clone(),
            init: init.clone(),
            modifiers: Modifiers::default(),
            with_events: false,
            array_bounds: array_bounds.clone(),
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

fn pascal_method_name_text(pair: Pair<Rule>) -> String {
    pair.into_inner()
        .next()
        .map(|inner| inner.as_str().to_string())
        .unwrap_or_default()
}

fn walk_class_constructor_sig(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut params = Vec::new();
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::method_sig_body => {
                for sp in p.into_inner() {
                    if sp.as_rule() == Rule::param_clause {
                        params = walk_param_clause(sp)?;
                    }
                }
            }
            Rule::procedure_body => body = walk_routine_body(p)?,
            _ => {}
        }
    }

    Ok(ClassMember::Constructor {
        name: None,
        params,
        body,
        base_args: None,
        initializer_target: vybe_ast::ConstructorInitializerTarget::Base,
        visibility: Visibility::Public,
    })
}

fn walk_class_method_sig(pair: Pair<Rule>, is_destructor: bool) -> Result<ClassMember, String> {
    let mut name = String::new();
    let mut params = Vec::new();
    let mut return_type: Option<String> = None;
    let mut modifiers = Modifiers::default();
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::method_sig_body => {
                for sp in p.into_inner() {
                    match sp.as_rule() {
                        Rule::method_name => name = pascal_method_name_text(sp),
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
            Rule::procedure_body | Rule::function_body => body = walk_routine_body(p)?,
            _ => {}
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
            body,
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
    let mut body = Vec::new();
    let mut modifiers = Modifiers {
        is_static: true,
        ..Default::default()
    };

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::method_sig_body => {
                for sp in p.into_inner() {
                    match sp.as_rule() {
                        Rule::method_name => name = pascal_method_name_text(sp),
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
            Rule::procedure_body | Rule::function_body => body = walk_routine_body(p)?,
            _ => {}
        }
    }

    if is_field {
        Ok(ClassMember::Field {
            name,
            init: type_hint.as_deref().and_then(default_field_init_for_type),
            type_hint,
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
                body,
                modifiers,
                handles: Vec::new(),
                is_async: false,
                is_generator: false,
                is_sub,
            },
        ))))
    }
}

fn walk_routine_body(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut body = Vec::new();
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::decl_section => walk_decl_section(p, &mut body)?,
            Rule::compound_statement => body.extend(walk_compound_statement(p)?),
            _ => {}
        }
    }
    Ok(body)
}

fn walk_class_property_decl(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let decl_text = pair.as_str().to_ascii_lowercase();
    let is_static = pair
        .as_str()
        .trim_start()
        .to_lowercase()
        .starts_with("class");
    let mut name = String::new();
    let mut type_hint: Option<String> = None;
    let mut getter: Option<Vec<Statement>> = None;
    let mut setter: Option<PropertySetter> = None;
    let mut modifiers = Modifiers {
        is_static,
        ..Default::default()
    };
    if decl_text.contains("; default") {
        modifiers
            .decorators
            .push(Expression::string("__pascal_default_property"));
    }

    for p in pair.into_inner() {
        if p.as_rule() == Rule::property_def {
            for sp in p.into_inner() {
                match sp.as_rule() {
                    Rule::identifier => {
                        if name.is_empty() {
                            name = sp.as_str().to_string();
                        }
                    }
                    Rule::property_index => {
                        modifiers
                            .decorators
                            .push(Expression::string("__pascal_indexed_property"));
                    }
                    Rule::type_ref => type_hint = Some(type_ref_to_string(&sp)),
                    Rule::property_specifiers => {
                        for spec in sp.into_inner() {
                            match spec.as_rule() {
                                Rule::property_read => {
                                    let getter_name = property_accessor_name(spec);
                                    getter = Some(vec![Statement::new(StmtKind::Return(Some(
                                        property_read_accessor(getter_name, is_static),
                                    )))]);
                                }
                                Rule::property_write => {
                                    let setter_name = property_accessor_name(spec);
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
                                        body: vec![property_write_accessor(
                                            setter_name,
                                            is_static,
                                            Expression::ident("value"),
                                        )],
                                    });
                                }
                                Rule::property_default => {
                                    modifiers
                                        .decorators
                                        .push(Expression::string("__pascal_default_property"));
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

fn property_accessor_name(spec: Pair<Rule>) -> String {
    spec.into_inner()
        .next()
        .map(|p| p.as_str().trim().to_string())
        .unwrap_or_default()
}

fn property_read_accessor(name: String, is_static: bool) -> Expression {
    if !name.contains('.') && name.get(..3).is_some_and(|p| p.eq_ignore_ascii_case("Get")) {
        return property_accessor_call(name, is_static, Vec::new());
    }
    if !name.contains('.') {
        return property_accessor_call(name, is_static, Vec::new());
    }
    property_accessor_member(name, is_static)
}

fn property_write_accessor(name: String, is_static: bool, value: Expression) -> Statement {
    if !name.contains('.') && name.get(..3).is_some_and(|p| p.eq_ignore_ascii_case("Set")) {
        return Statement::new(StmtKind::Expr(property_accessor_call(
            name,
            is_static,
            vec![Argument::positional(value)],
        )));
    }
    if !name.contains('.') {
        return Statement::new(StmtKind::Expr(property_accessor_call(
            name,
            is_static,
            vec![Argument::positional(value)],
        )));
    }
    Statement::new(StmtKind::Assign {
        targets: vec![property_accessor_member(name, is_static)],
        value,
    })
}

fn property_accessor_member(name: String, is_static: bool) -> Expression {
    let mut parts = name.split('.');
    let first = parts.next().unwrap_or_default();
    let _ = is_static;
    let mut expr = Expression::ident(first);
    for part in parts {
        expr = Expression::new(ExprKind::Member {
            object: Box::new(expr),
            field: part.to_string(),
            null_safe: false,
        });
    }
    expr
}

fn property_accessor_call(name: String, is_static: bool, args: Vec<Argument>) -> Expression {
    let callee = if is_static {
        Expression::ident(&name)
    } else {
        Expression::new(ExprKind::Member {
            object: Box::new(Expression::new(ExprKind::This)),
            field: name,
            null_safe: false,
        })
    };
    Expression::new(ExprKind::Call {
        callee: Box::new(callee),
        args,
        optional: false,
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
                        Rule::method_name => name = pascal_method_name_text(sp),
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
            name: None,
            params,
            body: Vec::new(),
            base_args: None,
            initializer_target: vybe_ast::ConstructorInitializerTarget::Base,
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
                    Rule::method_name => name = pascal_method_name_text(sp),
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
                    Rule::method_name => name = format!("operator_{}", pascal_method_name_text(sp)),
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
                "static" => modifiers.is_static = true,
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

fn walk_class_var_decl_impl(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut class_name = String::new();
    let mut field_name = String::new();
    let mut type_hint = None;
    let mut init = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::qualified_type_name => class_name = p.as_str().to_string(),
            Rule::identifier => field_name = p.as_str().to_string(),
            Rule::type_ref => type_hint = Some(type_ref_to_string(&p)),
            Rule::var_init => {
                for inner in p.into_inner() {
                    if inner.as_rule() == Rule::expression {
                        init = Some(walk_expression(inner)?);
                    }
                }
            }
            _ => {}
        }
    }

    Ok(Statement::with_span(
        StmtKind::ClassDecl {
            name: class_name,
            parents: Vec::new(),
            interfaces: Vec::new(),
            members: vec![ClassMember::Field {
                name: field_name,
                init: init.or_else(|| type_hint.as_deref().and_then(default_field_init_for_type)),
                type_hint,
                modifiers: Modifiers {
                    is_static: true,
                    ..Default::default()
                },
                with_events: false,
                array_bounds: None,
            }],
            modifiers: ClassModifiers::default(),
            decorators: Vec::new(),
        },
        span,
    ))
}

fn walk_method_impl_proc(pair: Pair<Rule>, span: Span) -> Result<Statement, String> {
    let mut class_name = String::new();
    let mut method_name = String::new();
    let mut params = Vec::new();
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => {
                if class_name.is_empty() {
                    class_name = p.as_str().to_string();
                }
            }
            Rule::method_name => method_name = pascal_method_name_text(p),
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

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => {
                if class_name.is_empty() {
                    class_name = p.as_str().to_string();
                }
            }
            Rule::method_name => method_name = pascal_method_name_text(p),
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

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => {
                if class_name.is_empty() {
                    class_name = p.as_str().to_string();
                }
            }
            Rule::method_name => method_name = pascal_method_name_text(p),
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
        Rule::label_statement => walk_label_statement(inner)?,
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
        Rule::goto_statement => walk_goto_statement(inner)?,
        Rule::inherited_statement => walk_inherited_statement(inner)?,
        Rule::assign_or_call_statement => walk_assign_or_call(inner)?,
        Rule::empty_statement => StmtKind::Empty,
        other => return Err(format!("Unexpected statement rule: {:?}", other)),
    };

    Ok(Statement::with_span(kind, span))
}

fn walk_label_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let name = pair
        .into_inner()
        .find(|p| p.as_rule() == Rule::identifier)
        .map(|p| p.as_str().to_string())
        .ok_or_else(|| "label_statement: missing label".to_string())?;
    Ok(StmtKind::Label(name))
}

fn walk_goto_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let target = pair
        .into_inner()
        .find(|p| p.as_rule() == Rule::identifier)
        .map(|p| p.as_str().to_string())
        .ok_or_else(|| "goto_statement: missing target".to_string())?;
    Ok(StmtKind::GoTo(target))
}

fn walk_inline_var_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut names = Vec::new();
    let mut type_hint: Option<String> = None;
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
                let _ = extract_array_bounds(&p)?;
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
        declarations: build_var_declarators(names, type_hint, init, None),
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
    let for_stmt = Statement::new(StmtKind::For {
        init: Some(Box::new(init)),
        cond: Some(cond),
        update: Some(update_assign),
        body: flatten_stmt(body_stmt),
    });
    let final_value = if use_char_ordinal_loop {
        end_expr
    } else {
        end_expr
    };
    let restore_stmt = Statement::new(StmtKind::Assign {
        targets: vec![Expression::ident(&var_name)],
        value: final_value,
    });
    Ok(StmtKind::Block(vec![for_stmt, restore_stmt]))
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
    let iter_expr = pascal_for_in_iter(walk_expression(parts.remove(0))?);
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

fn pascal_for_in_iter(expr: Expression) -> Expression {
    if let ExprKind::Lit(Literal::Str(value)) = &expr.kind {
        return Expression::new(ExprKind::Array(
            value
                .chars()
                .map(|ch| ArrayElement {
                    key: None,
                    value: Expression::new(ExprKind::Lit(Literal::Str(ch.to_string()))),
                    spread: false,
                    by_ref: false,
                })
                .collect(),
        ));
    }
    expr
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

    if let Some(name) = &var_name {
        rewrite_bare_raise_in_catch(&mut body, name);
    }

    Ok(CatchClause {
        types: vec![type_name],
        var_name,
        stack_var: None,
        body,
        when_clause: None,
    })
}

fn rewrite_bare_raise_in_catch(body: &mut [Statement], var_name: &str) {
    for stmt in body {
        match &mut stmt.kind {
            StmtKind::Throw { expr, .. } if expr.is_none() => {
                *expr = Some(Expression::ident(var_name));
            }
            StmtKind::Block(stmts) => rewrite_bare_raise_in_catch(stmts, var_name),
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                rewrite_bare_raise_in_catch(then_body, var_name);
                for (_, body) in elifs {
                    rewrite_bare_raise_in_catch(body, var_name);
                }
                if let Some(body) = else_body {
                    rewrite_bare_raise_in_catch(body, var_name);
                }
            }
            StmtKind::While { body, .. }
            | StmtKind::DoWhile { body, .. }
            | StmtKind::For { body, .. }
            | StmtKind::ForIn { body, .. } => rewrite_bare_raise_in_catch(body, var_name),
            StmtKind::Try {
                body,
                catches,
                finally,
                ..
            } => {
                rewrite_bare_raise_in_catch(body, var_name);
                for catch in catches {
                    if catch.var_name.is_none() {
                        rewrite_bare_raise_in_catch(&mut catch.body, var_name);
                    }
                }
                if let Some(finally) = finally {
                    rewrite_bare_raise_in_catch(finally, var_name);
                }
            }
            _ => {}
        }
    }
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

fn walk_halt_statement(_pair: Pair<Rule>) -> Result<StmtKind, String> {
    Ok(StmtKind::Return(None))
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
                if name.eq_ignore_ascii_case("Inc") && (args.len() == 1 || args.len() == 2) {
                    let target = args[0].value.clone();
                    return Ok(StmtKind::Assign {
                        targets: vec![target.clone()],
                        value: Expression::new(ExprKind::Binary {
                            op: BinOp::Add,
                            left: Box::new(target),
                            right: Box::new(
                                args.get(1)
                                    .map(|arg| arg.value.clone())
                                    .unwrap_or_else(|| Expression::int(1)),
                            ),
                        }),
                    });
                }
                if name.eq_ignore_ascii_case("Dec") && (args.len() == 1 || args.len() == 2) {
                    let target = args[0].value.clone();
                    return Ok(StmtKind::Assign {
                        targets: vec![target.clone()],
                        value: Expression::new(ExprKind::Binary {
                            op: BinOp::Sub,
                            left: Box::new(target),
                            right: Box::new(
                                args.get(1)
                                    .map(|arg| arg.value.clone())
                                    .unwrap_or_else(|| Expression::int(1)),
                            ),
                        }),
                    });
                }
                if name.eq_ignore_ascii_case("Str") && args.len() == 2 {
                    return Ok(StmtKind::Assign {
                        targets: vec![args[1].value.clone()],
                        value: pascal_str_value(args[0].value.clone()),
                    });
                }
                if name.eq_ignore_ascii_case("Val") && args.len() == 3 {
                    let target = args[1].value.clone();
                    let invalid = pascal_val_invalid_expr(args[0].value.clone());
                    return Ok(StmtKind::Block(vec![
                        Statement::new(StmtKind::Assign {
                            targets: vec![target.clone()],
                            value: Expression::new(ExprKind::Ternary {
                                cond: Box::new(invalid.clone()),
                                then: Box::new(Expression::new(ExprKind::Lit(Literal::Int(0)))),
                                else_: Box::new(Expression::new(ExprKind::Call {
                                    callee: Box::new(Expression::ident("parseInt")),
                                    args: vec![Argument::positional(args[0].value.clone())],
                                    optional: false,
                                })),
                            }),
                        }),
                        Statement::new(StmtKind::Assign {
                            targets: vec![args[2].value.clone()],
                            value: Expression::new(ExprKind::Ternary {
                                cond: Box::new(invalid),
                                then: Box::new(Expression::new(ExprKind::Lit(Literal::Int(1)))),
                                else_: Box::new(Expression::new(ExprKind::Lit(Literal::Int(0)))),
                            }),
                        }),
                    ]));
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
            ExprKind::Ident(_) | ExprKind::Member { .. } | ExprKind::Index { .. } => Expression::with_span(
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
            if let Some(source) = pascal_str_to_int_arg(&value) {
                return Ok(StmtKind::If {
                    cond: pascal_val_invalid_expr(source.clone()),
                    then_body: vec![Statement::new(StmtKind::Throw {
                        expr: Some(Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Member {
                                object: Box::new(Expression::ident("Exception")),
                                field: "Create".to_string(),
                                null_safe: false,
                            })),
                            args: vec![Argument::positional(Expression::new(ExprKind::Lit(
                                Literal::Str("invalid integer".to_string()),
                            )))],
                            optional: false,
                        })),
                        cause: None,
                    })],
                    elifs: Vec::new(),
                    else_body: Some(vec![Statement::new(StmtKind::Assign {
                        targets: vec![target],
                        value: Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident("parseInt")),
                            args: vec![Argument::positional(source)],
                            optional: false,
                        }),
                    })]),
                });
            }
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
            "><" => BinOp::BitXor,
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
            } else if let Some(rest) = s.strip_prefix('%') {
                Ok(ExprKind::Lit(Literal::Int(
                    i64::from_str_radix(rest, 2).unwrap_or(0),
                )))
            } else if let Some(rest) = s.strip_prefix('&') {
                Ok(ExprKind::Lit(Literal::Int(
                    i64::from_str_radix(rest, 8).unwrap_or(0),
                )))
            } else {
                Ok(ExprKind::Lit(Literal::Int(
                    s.replace('_', "").parse().unwrap_or(0),
                )))
            }
        }
        Rule::real_literal => Ok(ExprKind::Lit(Literal::Float(
            pair.as_str().parse().unwrap_or(0.0),
        ))),
        Rule::string_literal => {
            let raw = pair.as_str();
            // Strip surrounding quotes and unescape Pascal/Delphi spellings.
            let inner = &raw[1..raw.len() - 1];
            let value = if raw.starts_with('"') {
                inner.replace("\\\"", "\"")
            } else {
                inner.replace("''", "'")
            };
            Ok(ExprKind::Lit(Literal::Str(value)))
        }
        Rule::char_literal => {
            // #65 → 'A', #13#10 → "\r\n"
            let s = pair.as_str();
            let decoded: String = s
                .split('#')
                .filter(|part| !part.is_empty())
                .filter_map(|part| part.parse::<u32>().ok())
                .filter_map(char::from_u32)
                .collect();
            let mut chars = decoded.chars();
            if let (Some(ch), None) = (chars.next(), chars.next()) {
                Ok(ExprKind::Lit(Literal::Char(ch)))
            } else {
                Ok(ExprKind::Lit(Literal::Str(decoded)))
            }
        }
        Rule::identifier => Ok(pascal_bare_identifier_expr(pair.as_str()).kind),

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
            return Ok(pascal_bare_identifier_expr(src).kind);
        }
    };

    match inner.as_rule() {
        Rule::int_literal => walk_expr_kind(inner),
        Rule::real_literal => walk_expr_kind(inner),
        Rule::string_literal => walk_expr_kind(inner),
        Rule::char_literal => walk_expr_kind(inner),
        Rule::identifier => Ok(pascal_bare_identifier_expr(inner.as_str()).kind),
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
                Rule::method_name => ident = pascal_method_name_text(p.clone()),
                Rule::arg_list => arg_list = Some(p.clone()),
                _ => {}
            }
        }

        if ident.eq_ignore_ascii_case("ClassName") || ident.eq_ignore_ascii_case("ClassType") {
            if arg_list.is_none() {
                return Ok(Expression::new(ExprKind::Member {
                    object: Box::new(expr),
                    field: if ident.eq_ignore_ascii_case("ClassName") {
                        "__pascal_class_name".to_string()
                    } else {
                        "__pascal_class_type".to_string()
                    },
                    null_safe: false,
                }));
            }
        }

        if ident.eq_ignore_ascii_case("InheritsFrom") {
            if let Some(al) = arg_list.clone() {
                let args = walk_arg_list(al)?;
                if args.len() == 1 {
                    return Ok(Expression::new(ExprKind::IsType {
                        expr: Box::new(expr),
                        type_name: match &args[0].value.kind {
                            ExprKind::Ident(name) => name.clone(),
                            ExprKind::Lit(Literal::Str(name)) => name.clone(),
                            _ => return Ok(Expression::new(ExprKind::Lit(Literal::Bool(false)))),
                        },
                    }));
                }
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
            if let Some(date_expr) = lower_pascal_datetime_builtin(name, &args) {
                return Ok(date_expr);
            }
            if name.eq_ignore_ascii_case("FormatDateTime") && args.len() == 2 {
                if let Some(formatted) =
                    pascal_format_datetime_expr(&args[0].value, args[1].value.clone())
                {
                    return Ok(formatted);
                }
            }
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
            if name.eq_ignore_ascii_case("Odd") && args.len() == 1 {
                return Ok(pascal_odd_expr(args[0].value.clone()));
            }
            if name.eq_ignore_ascii_case("Frac") && args.len() == 1 {
                return Ok(pascal_frac_expr(args[0].value.clone()));
            }
            if name.eq_ignore_ascii_case("IntToStr") && args.len() == 1 {
                let value = args[0].value.clone();
                return Ok(Expression::new(ExprKind::Binary {
                    op: BinOp::Add,
                    left: Box::new(value),
                    right: Box::new(Expression::new(ExprKind::Lit(Literal::Str(String::new())))),
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

fn pascal_bare_identifier_expr(name: &str) -> Expression {
    if name.eq_ignore_ascii_case("Now") {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("Now")),
            args: Vec::new(),
            optional: false,
        })
    } else if name.eq_ignore_ascii_case("MinInt") {
        Expression::new(ExprKind::Lit(Literal::Int(i32::MIN as i64)))
    } else {
        Expression::ident(name)
    }
}

fn pascal_date_component_call(helper: &str, date: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(helper)),
        args: vec![Argument::positional(date)],
        optional: false,
    })
}

fn call_expr(name: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(name)),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

fn int_expr(value: i64) -> Expression {
    Expression::new(ExprKind::Lit(Literal::Int(value)))
}

fn str_expr(value: &str) -> Expression {
    Expression::new(ExprKind::Lit(Literal::Str(value.to_string())))
}

fn bin_expr(op: BinOp, left: Expression, right: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right),
    })
}

fn ternary_expr(cond: Expression, then: Expression, else_: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(cond),
        then: Box::new(then),
        else_: Box::new(else_),
    })
}

fn pascal_date_month_expr(date: Expression) -> Expression {
    bin_expr(
        BinOp::Add,
        pascal_date_component_call("__pascal_date_month", date),
        int_expr(1),
    )
}

fn assign_expr(target: Expression, value: Expression) -> Expression {
    Expression::new(ExprKind::Assign {
        target: Box::new(target),
        value: Box::new(value),
    })
}

fn pascal_days_between_expr(left: Expression, right: Expression) -> Expression {
    pascal_abs_div_expr(left, right, 86_400_000)
}

fn pascal_abs_div_expr(left: Expression, right: Expression, divisor: i64) -> Expression {
    call_expr(
        "__pascal_abs",
        vec![bin_expr(
            BinOp::Div,
            bin_expr(BinOp::Sub, left, right),
            int_expr(divisor),
        )],
    )
}

fn pascal_compare_expr(left: Expression, right: Expression) -> Expression {
    let diff = bin_expr(BinOp::Sub, left.clone(), right.clone());
    ternary_expr(
        bin_expr(BinOp::Eq, diff, int_expr(0)),
        int_expr(0),
        ternary_expr(bin_expr(BinOp::Lt, left, right), int_expr(-1), int_expr(1)),
    )
}

fn pascal_parse_slash_date_literal(expr: &Expression) -> Option<Expression> {
    let ExprKind::Lit(Literal::Str(text)) = &expr.kind else {
        return None;
    };
    let mut parts = text.split('/');
    let month = parts.next()?.trim().parse::<i64>().ok()?;
    let day = parts.next()?.trim().parse::<i64>().ok()?;
    let year = parts.next()?.trim().parse::<i64>().ok()?;
    if parts.next().is_some() {
        return None;
    }
    Some(call_expr(
        "__pascal_date_utc",
        vec![int_expr(year), int_expr(month - 1), int_expr(day)],
    ))
}

fn pascal_parse_time_literal(expr: &Expression) -> Option<Expression> {
    let ExprKind::Lit(Literal::Str(text)) = &expr.kind else {
        return None;
    };
    let mut parts = text.split(':');
    let hour = parts.next()?.trim().parse::<i64>().ok()?;
    let minute = parts.next()?.trim().parse::<i64>().ok()?;
    let second = parts
        .next()
        .map(|part| part.trim().parse::<i64>().ok())
        .unwrap_or(Some(0))?;
    if parts.next().is_some() {
        return None;
    }
    Some(call_expr(
        "__pascal_date_utc",
        vec![
            int_expr(1970),
            int_expr(0),
            int_expr(1),
            int_expr(hour),
            int_expr(minute),
            int_expr(second),
            int_expr(0),
        ],
    ))
}

fn lower_pascal_datetime_builtin(name: &str, args: &[Argument]) -> Option<Expression> {
    let arg = |idx: usize| args.get(idx).map(|arg| arg.value.clone());
    match (name.to_lowercase().as_str(), args.len()) {
        ("encodedate", 3) => Some(call_expr(
            "__pascal_date_utc",
            vec![arg(0)?, bin_expr(BinOp::Sub, arg(1)?, int_expr(1)), arg(2)?],
        )),
        ("encodetime", 4) => Some(call_expr(
            "__pascal_date_utc",
            vec![
                int_expr(1970),
                int_expr(0),
                int_expr(1),
                arg(0)?,
                arg(1)?,
                arg(2)?,
                arg(3)?,
            ],
        )),
        ("decodedate", 4) => Some(Expression::new(ExprKind::Sequence(vec![
            assign_expr(
                arg(1)?,
                pascal_date_component_call("__pascal_date_year", arg(0)?),
            ),
            assign_expr(arg(2)?, pascal_date_month_expr(arg(0)?)),
            assign_expr(
                arg(3)?,
                pascal_date_component_call("__pascal_date_day", arg(0)?),
            ),
            int_expr(0),
        ]))),
        ("decodetime", 5) => Some(Expression::new(ExprKind::Sequence(vec![
            assign_expr(
                arg(1)?,
                pascal_date_component_call("__pascal_date_hour", arg(0)?),
            ),
            assign_expr(
                arg(2)?,
                pascal_date_component_call("__pascal_date_minute", arg(0)?),
            ),
            assign_expr(
                arg(3)?,
                pascal_date_component_call("__pascal_date_second", arg(0)?),
            ),
            assign_expr(
                arg(4)?,
                pascal_date_component_call("__pascal_date_millisecond", arg(0)?),
            ),
            int_expr(0),
        ]))),
        ("dayof", 1) => Some(pascal_date_component_call("__pascal_date_day", arg(0)?)),
        ("monthof", 1) => Some(pascal_date_month_expr(arg(0)?)),
        ("yearof", 1) => Some(pascal_date_component_call("__pascal_date_year", arg(0)?)),
        ("hourof", 1) => Some(pascal_date_component_call("__pascal_date_hour", arg(0)?)),
        ("minuteof", 1) => Some(pascal_date_component_call("__pascal_date_minute", arg(0)?)),
        ("secondof", 1) => Some(pascal_date_component_call("__pascal_date_second", arg(0)?)),
        ("dayofweek", 1) => {
            let dow = pascal_date_component_call("__pascal_date_weekday", arg(0)?);
            Some(ternary_expr(
                bin_expr(BinOp::Eq, dow.clone(), int_expr(0)),
                int_expr(7),
                dow,
            ))
        }
        ("incday", 2) => Some(bin_expr(
            BinOp::Add,
            arg(0)?,
            bin_expr(BinOp::Mul, arg(1)?, int_expr(86_400_000)),
        )),
        ("inchour", 2) => Some(bin_expr(
            BinOp::Add,
            arg(0)?,
            bin_expr(BinOp::Mul, arg(1)?, int_expr(3_600_000)),
        )),
        ("incminute", 2) => Some(bin_expr(
            BinOp::Add,
            arg(0)?,
            bin_expr(BinOp::Mul, arg(1)?, int_expr(60_000)),
        )),
        ("incmonth", 2) => {
            let date = arg(0)?;
            Some(call_expr(
                "__pascal_date_utc",
                vec![
                    pascal_date_component_call("__pascal_date_year", date.clone()),
                    bin_expr(
                        BinOp::Add,
                        pascal_date_component_call("__pascal_date_month", date.clone()),
                        arg(1)?,
                    ),
                    call_expr(
                        "Min",
                        vec![
                            pascal_date_component_call("__pascal_date_day", date),
                            int_expr(28),
                        ],
                    ),
                ],
            ))
        }
        ("daysbetween", 2) => Some(pascal_days_between_expr(arg(0)?, arg(1)?)),
        ("hoursbetween", 2) => Some(pascal_abs_div_expr(arg(0)?, arg(1)?, 3_600_000)),
        ("minutesbetween", 2) => Some(pascal_abs_div_expr(arg(0)?, arg(1)?, 60_000)),
        ("comparedate", 2) => {
            let diff = pascal_days_between_expr(arg(0)?, arg(1)?);
            Some(ternary_expr(
                bin_expr(BinOp::Eq, diff.clone(), int_expr(0)),
                int_expr(0),
                ternary_expr(
                    bin_expr(BinOp::Lt, arg(0)?, arg(1)?),
                    int_expr(-1),
                    int_expr(1),
                ),
            ))
        }
        ("comparetime", 2) => Some(pascal_compare_expr(arg(0)?, arg(1)?)),
        ("samedate", 2) => Some(bin_expr(
            BinOp::Eq,
            pascal_days_between_expr(arg(0)?, arg(1)?),
            int_expr(0),
        )),
        ("strtodate", 1) => pascal_parse_slash_date_literal(&arg(0)?),
        ("strtotime", 1) => pascal_parse_time_literal(&arg(0)?),
        ("datetostr", 1) => pascal_format_datetime_expr(&str_expr("m/d/yyyy"), arg(0)?),
        ("timetostr", 1) => {
            let date = arg(0)?;
            let hour = pascal_date_component_call("__pascal_date_hour", date.clone());
            let hour_mod = bin_expr(BinOp::Mod, hour.clone(), int_expr(12));
            let hour12 = ternary_expr(
                bin_expr(BinOp::Eq, hour_mod.clone(), int_expr(0)),
                int_expr(12),
                hour_mod,
            );
            Some(call_expr(
                "Format",
                vec![
                    str_expr("%d:%02d:%02d %s"),
                    hour12,
                    pascal_date_component_call("__pascal_date_minute", date.clone()),
                    pascal_date_component_call("__pascal_date_second", date),
                    ternary_expr(
                        bin_expr(BinOp::Lt, hour, int_expr(12)),
                        str_expr("AM"),
                        str_expr("PM"),
                    ),
                ],
            ))
        }
        _ => None,
    }
}

fn pascal_format_datetime_expr(format: &Expression, date: Expression) -> Option<Expression> {
    let ExprKind::Lit(Literal::Str(picture)) = &format.kind else {
        return None;
    };

    let mut sprintf = String::new();
    let mut args = vec![Argument::positional(Expression::new(ExprKind::Lit(
        Literal::Str(String::new()),
    )))];
    let mut chars = picture.chars().peekable();
    while let Some(ch) = chars.next() {
        if ch.is_ascii_alphabetic() {
            let mut token = String::from(ch);
            while chars
                .peek()
                .is_some_and(|next| next.eq_ignore_ascii_case(&ch))
            {
                token.push(chars.next().unwrap());
            }

            let lower = token.to_ascii_lowercase();
            let width = token.len().min(2);
            let spec = if width == 2 { "%02d" } else { "%d" };
            let component = match lower.as_str() {
                "yyyy" => {
                    sprintf.push_str("%04d");
                    Some(pascal_date_component_call(
                        "__pascal_date_year",
                        date.clone(),
                    ))
                }
                "yy" => {
                    sprintf.push_str("%02d");
                    Some(Expression::new(ExprKind::Binary {
                        op: BinOp::Mod,
                        left: Box::new(pascal_date_component_call(
                            "__pascal_date_year",
                            date.clone(),
                        )),
                        right: Box::new(Expression::new(ExprKind::Lit(Literal::Int(100)))),
                    }))
                }
                "m" | "mm" => {
                    sprintf.push_str(spec);
                    Some(pascal_date_month_expr(date.clone()))
                }
                "d" | "dd" => {
                    sprintf.push_str(spec);
                    Some(pascal_date_component_call(
                        "__pascal_date_day",
                        date.clone(),
                    ))
                }
                "h" | "hh" => {
                    sprintf.push_str(spec);
                    Some(pascal_date_component_call(
                        "__pascal_date_hour",
                        date.clone(),
                    ))
                }
                "n" | "nn" => {
                    sprintf.push_str(spec);
                    Some(pascal_date_component_call(
                        "__pascal_date_minute",
                        date.clone(),
                    ))
                }
                "s" | "ss" => {
                    sprintf.push_str(spec);
                    Some(pascal_date_component_call(
                        "__pascal_date_second",
                        date.clone(),
                    ))
                }
                _ => {
                    sprintf.push_str(&token.replace('%', "%%"));
                    None
                }
            };
            if let Some(component) = component {
                args.push(Argument::positional(component));
            }
        } else {
            if ch == '%' {
                sprintf.push('%');
            }
            sprintf.push(ch);
        }
    }

    args[0] = Argument::positional(Expression::new(ExprKind::Lit(Literal::Str(sprintf))));
    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("Format")),
        args,
        optional: false,
    }))
}

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
    let mut args = Vec::new();
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::format_arg => args.push(walk_format_arg(p)?),
            Rule::expression => args.push(Argument::positional(walk_expression(p)?)),
            _ => {}
        }
    }
    Ok(args)
}

fn walk_format_arg(pair: Pair<Rule>) -> Result<Argument, String> {
    let mut inner = pair
        .into_inner()
        .filter(|p| p.as_rule() == Rule::expression);
    let value = inner
        .next()
        .map(walk_expression)
        .transpose()?
        .ok_or_else(|| "format_arg: missing value".to_string())?;
    let width = inner.next().map(walk_expression).transpose()?;
    let precision = inner.next().map(walk_expression).transpose()?;

    if let Some(width) = width {
        let value = format_value_expr(value, width, precision);
        Ok(Argument::positional(value))
    } else {
        Ok(Argument::positional(value))
    }
}

fn format_value_expr(
    value: Expression,
    width: Expression,
    precision: Option<Expression>,
) -> Expression {
    let width = const_int_expr(&width).unwrap_or(0);
    let fmt = if let Some(precision) = precision {
        let precision = const_int_expr(&precision).unwrap_or(0);
        format!("%{}.{}f", width, precision)
    } else {
        format!("%{}s", width)
    };

    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("Format")),
        args: vec![
            Argument::positional(Expression::new(ExprKind::Lit(Literal::Str(fmt)))),
            Argument::positional(value),
        ],
        optional: false,
    })
}

fn pascal_str_value(value: Expression) -> Expression {
    if matches!(&value.kind, ExprKind::Call { callee, .. } if matches!(&callee.kind, ExprKind::Ident(name) if name.eq_ignore_ascii_case("Format")))
    {
        value
    } else {
        Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(value),
            right: Box::new(Expression::new(ExprKind::Lit(Literal::Str(String::new())))),
        })
    }
}

fn pascal_odd_expr(value: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::NotEq,
        left: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Mod,
            left: Box::new(value),
            right: Box::new(Expression::new(ExprKind::Lit(Literal::Int(2)))),
        })),
        right: Box::new(Expression::new(ExprKind::Lit(Literal::Int(0)))),
    })
}

fn pascal_frac_expr(value: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::Sub,
        left: Box::new(value.clone()),
        right: Box::new(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("Trunc")),
            args: vec![Argument::positional(value)],
            optional: false,
        })),
    })
}

fn pascal_val_invalid_expr(value: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("isNaN")),
        args: vec![Argument::positional(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("Number")),
            args: vec![Argument::positional(value)],
            optional: false,
        }))],
        optional: false,
    })
}

fn pascal_str_to_int_arg(expr: &Expression) -> Option<Expression> {
    if let ExprKind::Call { callee, args, .. } = &expr.kind {
        if matches!(&callee.kind, ExprKind::Ident(name) if name.eq_ignore_ascii_case("StrToInt"))
            && args.len() == 1
        {
            return Some(args[0].value.clone());
        }
    }
    None
}

fn const_int_expr(expr: &Expression) -> Option<i64> {
    match expr.kind {
        ExprKind::Lit(Literal::Int(value)) => Some(value),
        ExprKind::Lit(Literal::Float(value)) => Some(value as i64),
        _ => None,
    }
}

// ── Set literal ────────────────────────────────────────────────────────────

fn walk_set_literal(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let elements: Vec<ArrayElement> = pair
        .into_inner()
        .filter(|p| p.as_rule() == Rule::set_element)
        .map(|p| {
            let mut parts = p.into_inner();
            let first = parts.next().ok_or("Empty set element")?;
            let start = walk_expression(first)?;
            let (value, spread) = if let Some(end_pair) = parts.next() {
                (
                    Expression::new(ExprKind::Range {
                        start: Box::new(start),
                        end: Box::new(walk_expression(end_pair)?),
                        inclusive: true,
                    }),
                    true,
                )
            } else {
                (start, false)
            };
            Ok(ArrayElement {
                key: None,
                value,
                spread,
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
            Rule::builtin_type_keyword => type_name = p.as_str().to_string(),
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

        left = if op_str == "><" {
            let left_minus_right = Expression::new(ExprKind::Binary {
                op: BinOp::Sub,
                left: Box::new(left.clone()),
                right: Box::new(right.clone()),
            });
            let right_minus_left = Expression::new(ExprKind::Binary {
                op: BinOp::Sub,
                left: Box::new(right),
                right: Box::new(left),
            });
            Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(left_minus_right),
                right: Box::new(right_minus_left),
            })
        } else {
            Expression::new(ExprKind::Binary {
                op: bin_op,
                left: Box::new(left),
                right: Box::new(right),
            })
        };
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
    normalize_pascal_type_hint(pair.as_str().trim())
}

fn normalize_pascal_type_hint(type_hint: &str) -> String {
    let trimmed = type_hint.trim();
    let lower = trimmed.to_ascii_lowercase();
    if lower.starts_with("tarray<") && trimmed.ends_with('>') {
        let inner = trimmed["TArray<".len()..trimmed.len() - 1].trim();
        return format!("array of {}", normalize_pascal_type_hint(inner));
    }
    trimmed.to_string()
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

#[derive(Debug, Clone)]
struct PascalFileInfo {
    path_var: Option<String>,
    is_text: bool,
}

fn lower_pascal_file_io(body: &mut Vec<Statement>) {
    static NEXT_PASCAL_FILE_HANDLE: std::sync::atomic::AtomicI64 =
        std::sync::atomic::AtomicI64::new(10_000);
    let mut next_handle =
        NEXT_PASCAL_FILE_HANDLE.fetch_add(1_000, std::sync::atomic::Ordering::Relaxed);
    let mut scope = std::collections::HashMap::new();
    let mut aliases = std::collections::HashMap::new();
    lower_pascal_file_io_body(body, &mut next_handle, &mut scope, &mut aliases);
}

fn lower_pascal_file_io_body(
    body: &mut Vec<Statement>,
    next_handle: &mut i64,
    scope: &mut std::collections::HashMap<String, PascalFileInfo>,
    aliases: &mut std::collections::HashMap<String, String>,
) {
    let mut out = Vec::with_capacity(body.len());
    for mut stmt in std::mem::take(body) {
        if let StmtKind::VarDecl { declarations, kind } = &mut stmt.kind {
            let mut companions = Vec::new();
            for decl in declarations.iter_mut() {
                let (BindingPattern::Ident(name), Some(type_hint)) =
                    (&decl.pattern, decl.type_hint.as_deref())
                else {
                    continue;
                };
                if *kind == VarDeclKind::Const && decl.init.is_none() {
                    if is_pascal_file_type_hint(type_hint, aliases) {
                        aliases.insert(name.to_lowercase(), type_hint.to_string());
                    }
                    continue;
                }
                if !is_pascal_file_type_hint(type_hint, aliases) {
                    continue;
                }
                let is_text = is_pascal_text_file_type_hint(type_hint, aliases);
                let handle = *next_handle;
                *next_handle += 1;
                if decl.init.is_none() {
                    decl.init = Some(Expression::int(handle));
                }
                let path_var = format!("__pascal_file_path_{}", handle);
                scope.insert(
                    name.to_lowercase(),
                    PascalFileInfo {
                        path_var: Some(path_var.clone()),
                        is_text,
                    },
                );
                companions.push(VarDeclarator {
                    pattern: BindingPattern::Ident(path_var),
                    type_hint: Some("String".to_string()),
                    init: Some(Expression::string("")),
                    array_bounds: None,
                    with_events: false,
                });
            }
            declarations.extend(companions);
        }

        lower_pascal_file_io_stmt(&mut stmt, next_handle, scope, aliases);
        out.push(stmt);
    }
    *body = out;
}

fn is_pascal_file_type_hint(
    type_hint: &str,
    aliases: &std::collections::HashMap<String, String>,
) -> bool {
    let lower = normalize_pascal_type_hint(type_hint).to_ascii_lowercase();
    lower == "text"
        || lower == "textfile"
        || lower == "file"
        || lower.starts_with("file of ")
        || aliases
            .get(&lower)
            .is_some_and(|aliased| is_pascal_file_type_hint(aliased, aliases))
}

fn is_pascal_text_file_type_hint(
    type_hint: &str,
    aliases: &std::collections::HashMap<String, String>,
) -> bool {
    let lower = normalize_pascal_type_hint(type_hint).to_ascii_lowercase();
    lower == "text"
        || lower == "textfile"
        || aliases
            .get(&lower)
            .is_some_and(|aliased| is_pascal_text_file_type_hint(aliased, aliases))
}

fn lower_pascal_file_io_stmt(
    stmt: &mut Statement,
    next_handle: &mut i64,
    scope: &mut std::collections::HashMap<String, PascalFileInfo>,
    aliases: &mut std::collections::HashMap<String, String>,
) {
    let replacement = match &mut stmt.kind {
        StmtKind::Expr(expr) => lower_pascal_file_io_call_stmt(expr, scope),
        StmtKind::Assign { value, .. } => {
            lower_pascal_file_io_expr(value, scope);
            None
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            lower_pascal_file_io_expr(target, scope);
            lower_pascal_file_io_expr(value, scope);
            None
        }
        StmtKind::Block(stmts) => {
            let mut scoped = scope.clone();
            let mut scoped_aliases = aliases.clone();
            lower_pascal_file_io_body(stmts, next_handle, &mut scoped, &mut scoped_aliases);
            None
        }
        StmtKind::FunctionDecl { params, body, .. } => {
            let mut scoped = scope.clone();
            for param in params {
                if param
                    .type_hint
                    .as_deref()
                    .is_some_and(|hint| is_pascal_file_type_hint(hint, aliases))
                {
                    let is_text = param
                        .type_hint
                        .as_deref()
                        .is_some_and(|hint| is_pascal_text_file_type_hint(hint, aliases));
                    scoped.insert(
                        param.name.to_lowercase(),
                        PascalFileInfo {
                            path_var: None,
                            is_text,
                        },
                    );
                }
            }
            let mut scoped_aliases = aliases.clone();
            lower_pascal_file_io_body(body, next_handle, &mut scoped, &mut scoped_aliases);
            None
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            lower_pascal_file_io_expr(cond, scope);
            let mut scoped = scope.clone();
            let mut scoped_aliases = aliases.clone();
            lower_pascal_file_io_body(then_body, next_handle, &mut scoped, &mut scoped_aliases);
            for (cond, body) in elifs {
                lower_pascal_file_io_expr(cond, scope);
                let mut scoped = scope.clone();
                let mut scoped_aliases = aliases.clone();
                lower_pascal_file_io_body(body, next_handle, &mut scoped, &mut scoped_aliases);
            }
            if let Some(body) = else_body {
                let mut scoped = scope.clone();
                let mut scoped_aliases = aliases.clone();
                lower_pascal_file_io_body(body, next_handle, &mut scoped, &mut scoped_aliases);
            }
            None
        }
        StmtKind::While { cond, body, .. } => {
            lower_pascal_file_io_expr(cond, scope);
            let mut scoped = scope.clone();
            let mut scoped_aliases = aliases.clone();
            lower_pascal_file_io_body(body, next_handle, &mut scoped, &mut scoped_aliases);
            None
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                lower_pascal_file_io_stmt(init, next_handle, scope, aliases);
            }
            if let Some(cond) = cond {
                lower_pascal_file_io_expr(cond, scope);
            }
            if let Some(update) = update {
                lower_pascal_file_io_expr(update, scope);
            }
            let mut scoped = scope.clone();
            let mut scoped_aliases = aliases.clone();
            lower_pascal_file_io_body(body, next_handle, &mut scoped, &mut scoped_aliases);
            None
        }
        StmtKind::ForIn { iter, body, .. } => {
            lower_pascal_file_io_expr(iter, scope);
            let mut scoped = scope.clone();
            let mut scoped_aliases = aliases.clone();
            lower_pascal_file_io_body(body, next_handle, &mut scoped, &mut scoped_aliases);
            None
        }
        StmtKind::DoWhile { body, cond, .. } => {
            let mut scoped = scope.clone();
            let mut scoped_aliases = aliases.clone();
            lower_pascal_file_io_body(body, next_handle, &mut scoped, &mut scoped_aliases);
            lower_pascal_file_io_expr(cond, scope);
            None
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            lower_pascal_file_io_expr(expr, scope);
            for case in cases {
                for cond in &mut case.conditions {
                    lower_pascal_file_io_case_condition(cond, scope);
                }
                let mut scoped = scope.clone();
                let mut scoped_aliases = aliases.clone();
                lower_pascal_file_io_body(
                    &mut case.body,
                    next_handle,
                    &mut scoped,
                    &mut scoped_aliases,
                );
            }
            if let Some(body) = default {
                let mut scoped = scope.clone();
                let mut scoped_aliases = aliases.clone();
                lower_pascal_file_io_body(body, next_handle, &mut scoped, &mut scoped_aliases);
            }
            None
        }
        StmtKind::ClassDecl { members, .. } | StmtKind::StructDecl { members, .. } => {
            for member in members {
                lower_pascal_file_io_member(member, next_handle, scope, aliases);
            }
            None
        }
        _ => None,
    };
    if let Some(kind) = replacement {
        stmt.kind = kind;
    }
}

fn lower_pascal_file_io_member(
    member: &mut ClassMember,
    next_handle: &mut i64,
    scope: &std::collections::HashMap<String, PascalFileInfo>,
    aliases: &std::collections::HashMap<String, String>,
) {
    match member {
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            let mut scoped = scope.clone();
            let mut scoped_aliases = aliases.clone();
            lower_pascal_file_io_stmt(stmt, next_handle, &mut scoped, &mut scoped_aliases);
        }
        ClassMember::Constructor { params, body, .. } => {
            let mut scoped = scope.clone();
            for param in params {
                if param
                    .type_hint
                    .as_deref()
                    .is_some_and(|hint| is_pascal_file_type_hint(hint, aliases))
                {
                    let is_text = param
                        .type_hint
                        .as_deref()
                        .is_some_and(|hint| is_pascal_text_file_type_hint(hint, aliases));
                    scoped.insert(
                        param.name.to_lowercase(),
                        PascalFileInfo {
                            path_var: None,
                            is_text,
                        },
                    );
                }
            }
            let mut scoped_aliases = aliases.clone();
            lower_pascal_file_io_body(body, next_handle, &mut scoped, &mut scoped_aliases);
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(body) = getter {
                let mut scoped = scope.clone();
                let mut scoped_aliases = aliases.clone();
                lower_pascal_file_io_body(body, next_handle, &mut scoped, &mut scoped_aliases);
            }
            if let Some(setter) = setter {
                let mut scoped = scope.clone();
                if setter
                    .param
                    .type_hint
                    .as_deref()
                    .is_some_and(|hint| is_pascal_file_type_hint(hint, aliases))
                {
                    let is_text = setter
                        .param
                        .type_hint
                        .as_deref()
                        .is_some_and(|hint| is_pascal_text_file_type_hint(hint, aliases));
                    scoped.insert(
                        setter.param.name.to_lowercase(),
                        PascalFileInfo {
                            path_var: None,
                            is_text,
                        },
                    );
                }
                let mut scoped_aliases = aliases.clone();
                lower_pascal_file_io_body(
                    &mut setter.body,
                    next_handle,
                    &mut scoped,
                    &mut scoped_aliases,
                );
            }
        }
        _ => {}
    }
}

fn lower_pascal_file_io_case_condition(
    cond: &mut CaseCondition,
    scope: &std::collections::HashMap<String, PascalFileInfo>,
) {
    match cond {
        CaseCondition::Value(expr) | CaseCondition::Comparison { expr, .. } => {
            lower_pascal_file_io_expr(expr, scope)
        }
        CaseCondition::Range { from, to } => {
            lower_pascal_file_io_expr(from, scope);
            lower_pascal_file_io_expr(to, scope);
        }
    }
}

fn lower_pascal_file_io_call_stmt(
    expr: &mut Expression,
    scope: &std::collections::HashMap<String, PascalFileInfo>,
) -> Option<StmtKind> {
    let ExprKind::Call { callee, args, .. } = &mut expr.kind else {
        lower_pascal_file_io_expr(expr, scope);
        return None;
    };
    let ExprKind::Ident(name) = &callee.kind else {
        lower_pascal_file_io_expr(expr, scope);
        return None;
    };
    let lowered = name.to_ascii_lowercase();
    let file_expr = args.first().map(|arg| arg.value.clone())?;
    let info = pascal_file_info_for_expr(&file_expr, scope)?;

    match lowered.as_str() {
        "assign" | "assignfile" if args.len() >= 2 => {
            let path_var = info.path_var.as_ref()?;
            Some(StmtKind::Assign {
                targets: vec![Expression::ident(path_var)],
                value: args[1].value.clone(),
            })
        }
        "rewrite" => Some(StmtKind::OpenFile {
            path: pascal_file_path_expr(&info)?,
            mode: FileMode::Output,
            file_number: file_expr,
        }),
        "reset" => Some(StmtKind::OpenFile {
            path: pascal_file_path_expr(&info)?,
            mode: FileMode::Input,
            file_number: file_expr,
        }),
        "append" => Some(StmtKind::OpenFile {
            path: pascal_file_path_expr(&info)?,
            mode: FileMode::Append,
            file_number: file_expr,
        }),
        "close" | "closefile" => Some(StmtKind::CloseFile(Some(file_expr))),
        "erase" => Some(StmtKind::Expr(pascal_call(
            "__pascal_file_remove",
            vec![pascal_file_path_expr(&info)?],
        ))),
        "rename" if args.len() >= 2 => {
            let path_var = info.path_var.as_ref()?;
            let new_path = args[1].value.clone();
            Some(StmtKind::Block(vec![
                Statement::new(StmtKind::Expr(pascal_call(
                    "__pascal_file_rename",
                    vec![Expression::ident(path_var), new_path.clone()],
                ))),
                Statement::new(StmtKind::Assign {
                    targets: vec![Expression::ident(path_var)],
                    value: new_path,
                }),
            ]))
        }
        "writeln" => {
            if !info.is_text && args.len() > 2 {
                return Some(StmtKind::Block(
                    args.iter()
                        .skip(1)
                        .map(|arg| {
                            Statement::new(StmtKind::PrintFile {
                                file_number: file_expr.clone(),
                                items: vec![arg.value.clone()],
                            })
                        })
                        .collect(),
                ));
            }
            let items = args.iter().skip(1).map(|arg| arg.value.clone()).collect();
            Some(StmtKind::PrintFile {
                file_number: file_expr,
                items,
            })
        }
        "write" => {
            if !info.is_text && args.len() > 2 {
                return Some(StmtKind::Block(
                    args.iter()
                        .skip(1)
                        .map(|arg| {
                            Statement::new(StmtKind::PrintFile {
                                file_number: file_expr.clone(),
                                items: vec![arg.value.clone()],
                            })
                        })
                        .collect(),
                ));
            }
            let items = args.iter().skip(1).map(|arg| arg.value.clone()).collect();
            Some(StmtKind::PrintFile {
                file_number: file_expr,
                items,
            })
        }
        "readln" if args.len() == 1 => Some(StmtKind::Expr(pascal_call(
            "__pascal_file_readline",
            vec![file_expr],
        ))),
        "readln" if args.len() == 2 => {
            let target = args[1].value.clone();
            if let ExprKind::Ident(var) = &target.kind {
                Some(StmtKind::LineInput {
                    file_number: file_expr,
                    variable: var.clone(),
                })
            } else {
                Some(StmtKind::Assign {
                    targets: vec![target],
                    value: pascal_call("__pascal_file_readline", vec![file_expr]),
                })
            }
        }
        "read" if args.len() == 2 => {
            let target = args[1].value.clone();
            if !info.is_text {
                if let ExprKind::Ident(var) = &target.kind {
                    Some(StmtKind::InputFile {
                        file_number: file_expr,
                        variables: vec![Expression::ident(var)],
                    })
                } else {
                    Some(StmtKind::Assign {
                        targets: vec![target],
                        value: pascal_call("__pascal_file_readline", vec![file_expr]),
                    })
                }
            } else if let ExprKind::Ident(var) = &target.kind {
                Some(StmtKind::LineInput {
                    file_number: file_expr,
                    variable: var.clone(),
                })
            } else {
                Some(StmtKind::Assign {
                    targets: vec![target],
                    value: pascal_call("__pascal_file_readline", vec![file_expr]),
                })
            }
        }
        "read" if args.len() > 2 && !info.is_text => Some(StmtKind::Block(
            args.iter()
                .skip(1)
                .map(|arg| {
                    let target = arg.value.clone();
                    if let ExprKind::Ident(var) = &target.kind {
                        Statement::new(StmtKind::InputFile {
                            file_number: file_expr.clone(),
                            variables: vec![Expression::ident(var)],
                        })
                    } else {
                        Statement::new(StmtKind::Assign {
                            targets: vec![target],
                            value: pascal_call("__pascal_file_readline", vec![file_expr.clone()]),
                        })
                    }
                })
                .collect(),
        )),
        _ => {
            for arg in args {
                lower_pascal_file_io_expr(&mut arg.value, scope);
            }
            None
        }
    }
}

fn pascal_file_path_expr(info: &PascalFileInfo) -> Option<Expression> {
    info.path_var.as_deref().map(Expression::ident)
}

fn pascal_file_info_for_expr(
    expr: &Expression,
    scope: &std::collections::HashMap<String, PascalFileInfo>,
) -> Option<PascalFileInfo> {
    match &expr.kind {
        ExprKind::Ident(name) => scope.get(&name.to_lowercase()).cloned(),
        _ => None,
    }
}

fn pascal_call(name: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(name)),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

fn lower_pascal_file_io_expr(
    expr: &mut Expression,
    scope: &std::collections::HashMap<String, PascalFileInfo>,
) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            lower_pascal_file_io_expr(callee, scope);
            if let ExprKind::Ident(name) = &callee.kind {
                if name.eq_ignore_ascii_case("eof") && args.len() == 1 {
                    if pascal_file_info_for_expr(&args[0].value, scope).is_some() {
                        return;
                    }
                }
            }
            for arg in args {
                lower_pascal_file_io_expr(&mut arg.value, scope);
            }
        }
        ExprKind::Member { object, .. } => lower_pascal_file_io_expr(object, scope),
        ExprKind::Index { object, index, .. } => {
            lower_pascal_file_io_expr(object, scope);
            lower_pascal_file_io_expr(index, scope);
        }
        ExprKind::Binary { left, right, .. } => {
            lower_pascal_file_io_expr(left, scope);
            lower_pascal_file_io_expr(right, scope);
        }
        ExprKind::Unary { expr, .. } => lower_pascal_file_io_expr(expr, scope),
        ExprKind::Ternary { cond, then, else_ } => {
            lower_pascal_file_io_expr(cond, scope);
            lower_pascal_file_io_expr(then, scope);
            lower_pascal_file_io_expr(else_, scope);
        }
        ExprKind::Assign { target, value } => {
            lower_pascal_file_io_expr(target, scope);
            lower_pascal_file_io_expr(value, scope);
        }
        ExprKind::Array(items) => {
            for item in items {
                lower_pascal_file_io_expr(&mut item.value, scope);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                if let ObjectProperty::KeyValue { key, value } = prop {
                    lower_pascal_file_io_expr(key, scope);
                    lower_pascal_file_io_expr(value, scope);
                }
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) => {
            for item in items {
                lower_pascal_file_io_expr(item, scope);
            }
        }
        ExprKind::New { class, args } => {
            lower_pascal_file_io_expr(class, scope);
            for arg in args {
                lower_pascal_file_io_expr(&mut arg.value, scope);
            }
        }
        _ => {}
    }
}

/// Get first inner pair from a compound pair.
fn cv_first(pair: Pair<Rule>) -> Result<Pair<Rule>, String> {
    pair.into_inner()
        .next()
        .ok_or_else(|| "Expected inner pair".to_string())
}
