//! Call-expression compilation — `compile_call` (handles named calls,
//! method calls, super-calls, spread, dotted lookups) and
//! `compile_lambda`. This is the primary edit site for the inline
//! refactor (Phase G) where `wasm:js-*` imports get replaced by
//! inline WASM GC sequences.

use super::*;

fn python_is_identifier_literal(value: &str) -> bool {
    let mut chars = value.chars();
    match chars.next() {
        Some(ch) if ch == '_' || ch.is_alphabetic() => {}
        _ => return false,
    }
    chars.all(|ch| ch == '_' || ch.is_alphanumeric())
}

fn python_is_printable_literal(value: &str) -> bool {
    value.chars().all(|ch| !ch.is_control())
}

fn terminal_type_name(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => Some(name.rsplit('.').next().unwrap_or(name).to_string()),
        ExprKind::Member { field, .. } => Some(field.clone()),
        _ => None,
    }
}

fn strip_generic_suffix(name: &str) -> &str {
    let trimmed = name.trim();
    let angle = trimmed.find('<');
    let vb = trimmed.to_ascii_lowercase().find("(of");
    match (angle, vb) {
        (Some(a), Some(b)) => trimmed[..a.min(b)].trim(),
        (Some(a), None) => trimmed[..a].trim(),
        (None, Some(b)) => trimmed[..b].trim(),
        (None, None) => trimmed,
    }
}

fn extract_generic_type_name(name: &str) -> Option<String> {
    let start = name.find('<')?;
    let end = name.rfind('>')?;
    let inner = name[start + 1..end].trim();
    Some(inner.rsplit('.').next().unwrap_or(inner).trim().to_string())
}

fn dotnet_factory_return_type(callee: &Expression) -> Option<String> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let class_name = terminal_type_name(object)?;
    if class_name.eq_ignore_ascii_case("TimeSpan")
        && matches!(
            field.as_str(),
            "FromDays" | "FromHours" | "FromMinutes" | "FromSeconds" | "FromMilliseconds" | "Zero"
        )
    {
        return Some("TimeSpan".into());
    }
    if class_name.eq_ignore_ascii_case("DateTime")
        && matches!(field.as_str(), "Now" | "UtcNow" | "Today" | "Parse")
    {
        return Some("DateTime".into());
    }
    if class_name.eq_ignore_ascii_case("Convert") && field.eq_ignore_ascii_case("ToDateTime") {
        return Some("DateTime".into());
    }
    if class_name.eq_ignore_ascii_case("Guid")
        && matches!(field.as_str(), "Empty" | "NewGuid" | "Parse")
    {
        return Some("Guid".into());
    }
    if class_name.eq_ignore_ascii_case("Version") && matches!(field.as_str(), "Parse") {
        return Some("Version".into());
    }
    None
}

fn dotnet_static_member_return_type(expr: &Expression) -> Option<String> {
    let ExprKind::Member { object, field, .. } = &expr.kind else {
        return None;
    };
    let class_name = terminal_type_name(object)?;
    if class_name.eq_ignore_ascii_case("DateTime")
        && matches!(field.as_str(), "Now" | "UtcNow" | "Today")
    {
        return Some("DateTime".into());
    }
    if class_name.eq_ignore_ascii_case("TimeSpan") && field == "Zero" {
        return Some("TimeSpan".into());
    }
    if class_name.eq_ignore_ascii_case("Guid") && field == "Empty" {
        return Some("Guid".into());
    }
    if class_name.eq_ignore_ascii_case("Version") && field == "Parse" {
        return Some("Version".into());
    }
    None
}

fn js_dynamic_import_alias(module: &str) -> String {
    let suffix: String = module
        .chars()
        .map(|ch| if ch.is_ascii_alphanumeric() { ch } else { '_' })
        .collect();
    format!("__js_dynamic_import_{}", suffix)
}

#[allow(dead_code)]
#[derive(Clone, Copy)]
enum JsPromiseChainKind {
    Then,
    Catch,
    Finally,
}

#[allow(dead_code)]
const JS_PROMISE_CHAIN_INPUT: &str = "__js_promise_chain_input";
#[allow(dead_code)]
const JS_PROMISE_CHAIN_ON_FULFILLED: &str = "__js_promise_chain_on_fulfilled";
#[allow(dead_code)]
const JS_PROMISE_CHAIN_ON_REJECTED: &str = "__js_promise_chain_on_rejected";
#[allow(dead_code)]
const JS_PROMISE_CHAIN_ON_FINALLY: &str = "__js_promise_chain_on_finally";
#[allow(dead_code)]
const JS_PROMISE_CHAIN_ERROR: &str = "__js_promise_chain_error";

#[allow(dead_code)]
const JS_PROMISE_THEN_PARAMS: [&str; 3] = [
    JS_PROMISE_CHAIN_INPUT,
    JS_PROMISE_CHAIN_ON_FULFILLED,
    JS_PROMISE_CHAIN_ON_REJECTED,
];
#[allow(dead_code)]
const JS_PROMISE_CATCH_PARAMS: [&str; 2] = [JS_PROMISE_CHAIN_INPUT, JS_PROMISE_CHAIN_ON_REJECTED];
#[allow(dead_code)]
const JS_PROMISE_FINALLY_PARAMS: [&str; 2] = [JS_PROMISE_CHAIN_INPUT, JS_PROMISE_CHAIN_ON_FINALLY];

#[allow(dead_code)]
fn js_promise_chain_params(kind: JsPromiseChainKind) -> &'static [&'static str] {
    match kind {
        JsPromiseChainKind::Then => &JS_PROMISE_THEN_PARAMS,
        JsPromiseChainKind::Catch => &JS_PROMISE_CATCH_PARAMS,
        JsPromiseChainKind::Finally => &JS_PROMISE_FINALLY_PARAMS,
    }
}

#[allow(dead_code)]
fn js_ident(name: &str) -> Expression {
    Expression::new(ExprKind::Ident(name.to_string()))
}

#[allow(dead_code)]
fn js_await(expr: Expression) -> Expression {
    Expression::new(ExprKind::Await(Box::new(expr)))
}

#[allow(dead_code)]
fn js_call_ident(name: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(js_ident(name)),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

#[allow(dead_code)]
fn js_nullish_check(name: &str) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::Or,
        left: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::StrictEq,
            left: Box::new(js_ident(name)),
            right: Box::new(Expression::null()),
        })),
        right: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::StrictEq,
            left: Box::new(js_ident(name)),
            right: Box::new(Expression::new(ExprKind::Lit(Literal::Undefined))),
        })),
    })
}

#[allow(dead_code)]
fn js_promise_chain_body(kind: JsPromiseChainKind) -> Vec<Statement> {
    match kind {
        JsPromiseChainKind::Then => vec![Statement::new(StmtKind::Try {
            body: vec![Statement::new(StmtKind::If {
                cond: js_nullish_check(JS_PROMISE_CHAIN_ON_FULFILLED),
                then_body: vec![Statement::new(StmtKind::Return(Some(js_await(js_ident(
                    JS_PROMISE_CHAIN_INPUT,
                )))))],
                elifs: vec![],
                else_body: Some(vec![Statement::new(StmtKind::Return(Some(js_call_ident(
                    JS_PROMISE_CHAIN_ON_FULFILLED,
                    vec![js_await(js_ident(JS_PROMISE_CHAIN_INPUT))],
                ))))]),
            })],
            catches: vec![CatchClause {
                types: vec![],
                var_name: Some(JS_PROMISE_CHAIN_ERROR.to_string()),
                stack_var: None,
                body: vec![Statement::new(StmtKind::If {
                    cond: js_nullish_check(JS_PROMISE_CHAIN_ON_REJECTED),
                    then_body: vec![Statement::new(StmtKind::Throw {
                        expr: Some(js_ident(JS_PROMISE_CHAIN_ERROR)),
                        cause: None,
                    })],
                    elifs: vec![],
                    else_body: Some(vec![Statement::new(StmtKind::Return(Some(js_call_ident(
                        JS_PROMISE_CHAIN_ON_REJECTED,
                        vec![js_ident(JS_PROMISE_CHAIN_ERROR)],
                    ))))]),
                })],
                when_clause: None,
            }],
            else_body: None,
            finally: None,
        })],
        JsPromiseChainKind::Catch => vec![Statement::new(StmtKind::Try {
            body: vec![Statement::new(StmtKind::Return(Some(js_await(js_ident(
                JS_PROMISE_CHAIN_INPUT,
            )))))],
            catches: vec![CatchClause {
                types: vec![],
                var_name: Some(JS_PROMISE_CHAIN_ERROR.to_string()),
                stack_var: None,
                body: vec![Statement::new(StmtKind::If {
                    cond: js_nullish_check(JS_PROMISE_CHAIN_ON_REJECTED),
                    then_body: vec![Statement::new(StmtKind::Throw {
                        expr: Some(js_ident(JS_PROMISE_CHAIN_ERROR)),
                        cause: None,
                    })],
                    elifs: vec![],
                    else_body: Some(vec![Statement::new(StmtKind::Return(Some(js_call_ident(
                        JS_PROMISE_CHAIN_ON_REJECTED,
                        vec![js_ident(JS_PROMISE_CHAIN_ERROR)],
                    ))))]),
                })],
                when_clause: None,
            }],
            else_body: None,
            finally: None,
        })],
        JsPromiseChainKind::Finally => {
            let input_var = "__finally_val";
            vec![
                Statement::new(StmtKind::VarDecl {
                    kind: crate::ast::VarDeclKind::Let,
                    declarations: vec![crate::ast::VarDeclarator {
                        pattern: crate::ast::BindingPattern::Ident(input_var.to_string()),
                        init: Some(js_await(js_ident(JS_PROMISE_CHAIN_INPUT))),
                        type_hint: None,
                        array_bounds: None,
                        with_events: false,
                    }],
                }),
                Statement::new(StmtKind::If {
                    cond: js_nullish_check(JS_PROMISE_CHAIN_ON_FINALLY),
                    then_body: vec![],
                    elifs: vec![],
                    else_body: Some(vec![Statement::new(StmtKind::Expr(js_call_ident(
                        JS_PROMISE_CHAIN_ON_FINALLY,
                        vec![],
                    )))]),
                }),
                Statement::new(StmtKind::Return(Some(js_ident(input_var)))),
            ]
        }
    }
}

fn js_prefers_typed_member_dispatch(type_hint: &str) -> bool {
    matches!(
        type_hint,
        "string"
            | "weakmap"
            | "weakset"
            | "weakref"
            | "finalizationregistry"
            | "collator"
            | "numberformat"
            | "datetimeformat"
            | "listformat"
            | "pluralrules"
            | "relativetimeformat"
            | "segmenter"
            | "locale"
            | "displaynames"
            | "durationformat"
            | "textencoder"
            | "textdecoder"
    )
}

fn resolve_receiver_type_hint(compiler: &Compiler, recv: &Expression) -> Option<String> {
    match &recv.kind {
        ExprKind::Ident(local_name) => compiler
            .lookup_var_type_hint(local_name)
            .map(str::to_string)
            .or_else(|| {
                compiler
                    .scope()
                    .resolve_type_ci(local_name)
                    .map(|s| s.to_string())
            })
            .or_else(|| {
                let cn = compiler.canon(local_name);
                compiler.global_type_hints.get(&cn).cloned()
            })
            .or_else(|| compiler.is_class_static_field_type_hint(local_name))
            .map(|name| compiler.resolve_source_type_alias(&name)),
        ExprKind::Member { object, field, .. } => {
            if let Some(type_name) = dotnet_static_member_return_type(recv) {
                return Some(type_name);
            }
            let owner_is_self = matches!(&object.kind, ExprKind::This | ExprKind::Super)
                || matches!(&object.kind, ExprKind::Ident(n)
                if {
                    let cn = compiler.canon(n);
                    cn == compiler.profile.self_keyword
                        || cn == "me"
                        || cn == "this"
                        || cn == "mybase"
                });
            if owner_is_self {
                compiler.is_class_static_field_type_hint(field)
            } else if let ExprKind::Ident(owner) = &object.kind {
                let owner_name = owner.split('<').next().map(str::trim).unwrap_or(owner);
                let canon_field = compiler.canon(field);

                let mut owner_candidates = vec![owner_name.to_string()];
                let owner_canon = compiler.canon(owner_name);
                if owner_canon != owner_name {
                    owner_candidates.push(owner_canon);
                }

                for owner_key in owner_candidates {
                    let mut current = Some(owner_key.as_str());
                    while let Some(cn) = current {
                        if let Some(pc) = compiler.pending_classes.get(cn) {
                            if let Some(type_hint) = pc.static_field_types.get(&canon_field) {
                                return Some(compiler.resolve_source_type_alias(type_hint));
                            }
                            current = pc.parent.as_deref();
                        } else {
                            break;
                        }
                    }
                }
                None
            } else {
                None
            }
        }
        ExprKind::New { class, .. } => {
            terminal_type_name(class).map(|name| compiler.resolve_source_type_alias(&name))
        }
        ExprKind::Call { callee, args, .. } => {
            let arg_exprs: Vec<&Expression> = args.iter().map(|arg| &arg.value).collect();
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                if let Some(return_type) = compiler
                    .resolve_instance_method_overload(object, field, &arg_exprs, false)
                    .and_then(|overload| overload.return_type.clone())
                {
                    return Some(return_type);
                }

                if compiler.profile.namespaces.use_dotnet {
                    if let Some(receiver_type) = resolve_receiver_type_hint(compiler, object) {
                        if compiler
                            .resolve_pending_class_name_for_type_hint(&receiver_type)
                            .is_none()
                        {
                            let class_name = Compiler::normalize_type_hint(&receiver_type);
                            if let Some(return_type) = common::dotnet::surface()
                                .lookup_instance_method_return_type(
                                    &class_name,
                                    field,
                                    args.len() as u8,
                                )
                            {
                                return Some(return_type);
                            }
                        }
                    }
                }
            }

            let inferred = compiler
                .infer_function_return_type(callee)
                .or_else(|| dotnet_factory_return_type(callee))
                .or_else(|| match &callee.kind {
                    ExprKind::Ident(name) => {
                        let resolved = compiler.resolve_source_type_alias(name);
                        common::dotnet::surface()
                            .lookup_constructor(&resolved)
                            .map(|_| resolved)
                    }
                    ExprKind::Member { field, .. } => {
                        let resolved = compiler.resolve_source_type_alias(field);
                        common::dotnet::surface()
                            .lookup_constructor(&resolved)
                            .map(|_| resolved)
                    }
                    _ => None,
                });

            if inferred.is_some() {
                return inferred;
            }

            if compiler.profile.name == "go" {
                if let ExprKind::Member { object, field, .. } = &callee.kind {
                    if let Some(receiver_type) = resolve_receiver_type_hint(compiler, object) {
                        if let Some(class_name) =
                            compiler.resolve_pending_class_name_for_type_hint(&receiver_type)
                        {
                            if compiler
                                .pending_classes
                                .get(&class_name)
                                .is_some_and(|pending| {
                                    pending
                                        .instance_pointer_method_names
                                        .iter()
                                        .any(|name| compiler.canon(name) == compiler.canon(field))
                                })
                            {
                                return Some(receiver_type);
                            }
                        }
                    }
                }
            }

            None
        }
        _ => None,
    }
}

fn resolves_to_static_container_method(
    compiler: &Compiler,
    object: &Expression,
    field: &str,
) -> bool {
    let class_parts = compiler.flatten_member_chain(object);
    if class_parts.is_empty() {
        return false;
    }

    let head_name = class_parts.first().map(String::as_str).unwrap_or("");
    let full_canon = compiler.canon(&class_parts.join("."));
    let short_canon = compiler.canon(class_parts.last().map(String::as_str).unwrap_or(""));

    // If the head resolves to a known class (even if it's also in scope as a global variable —
    // Python and similar languages register class names as both), check whether the field is
    // actually a static method on that class before bailing out on the scope check.
    // This lets `C.add(2, 3)` dispatch as a static call when `add` is `@staticmethod`.
    let head_is_known_class = compiler.defined_classes.contains(full_canon.as_str())
        || compiler.defined_classes.contains(short_canon.as_str());

    if !head_is_known_class
        && (compiler.scope().resolve(head_name).is_some()
            || compiler.scope().resolve_ci(head_name).is_some()
            || compiler.lookup_var_type_hint(head_name).is_some())
    {
        return false;
    }

    [full_canon, short_canon]
        .into_iter()
        .any(|container_canon| {
            let method_canon = compiler.js_member_storage_name_for_class(&container_canon, field);
            // §15.7: statics are inherited through the constructor chain —
            // `Dog.describe()` where `describe` is declared on `Animal`
            // resolves to the parent's static. Walk ancestors so the call
            // site recognizes inherited statics (the runtime copies them
            // onto the child constructor).
            compiler.defined_classes.contains(&container_canon)
                && class_or_ancestor_has_static(compiler, &container_canon, &method_canon)
        })
}

/// Does `class_canon` or any ancestor declare a static method `method_canon`?
fn class_or_ancestor_has_static(
    compiler: &Compiler,
    class_canon: &str,
    method_canon: &str,
) -> bool {
    let mut current = Some(class_canon.to_string());
    let mut guard = 0;
    while let Some(name) = current {
        guard += 1;
        if guard > 64 {
            break;
        }
        let Some(pending) = compiler.pending_classes.get(name.as_str()) else {
            break;
        };
        if pending
            .static_method_names
            .iter()
            .any(|n| n == method_canon)
        {
            return true;
        }
        current = pending.parent.as_ref().map(|p| compiler.canon(p));
    }
    false
}

fn has_explicit_constructor_signature(compiler: &Compiler, class_name: &str) -> bool {
    compiler
        .constructor_signatures
        .get(class_name)
        .is_some_and(|signatures| !signatures.is_empty())
}

fn is_numeric_overload_type(type_hint: &str) -> bool {
    matches!(
        type_hint,
        "integer"
            | "int"
            | "int32"
            | "longint"
            | "long"
            | "int64"
            | "short"
            | "int16"
            | "uint"
            | "uint32"
            | "ulong"
            | "uint64"
            | "ushort"
            | "uint16"
            | "byte"
            | "sbyte"
            | "real"
            | "double"
            | "float"
            | "single"
            | "decimal"
    )
}

fn resolve_go_pending_instance_method_owner(
    compiler: &Compiler,
    object: &Expression,
    field: &str,
) -> Option<String> {
    if compiler.profile.name != "go" {
        return None;
    }
    let receiver_type = resolve_receiver_type_hint(compiler, object)?;
    let class_name = compiler.resolve_pending_class_name_for_type_hint(&receiver_type)?;
    let pending = compiler.pending_classes.get(&class_name)?;
    let method_key = compiler.canon(field);
    pending
        .instance_member_names
        .iter()
        .any(|name| compiler.canon(name) == method_key)
        .then_some(class_name)
}

impl Compiler {
    fn emit_fortran_member_receiver_writeback(&mut self, object: &Expression, receiver_slot: u16) {
        if self.profile.name != "fortran" {
            return;
        }
        let result_slot = self.define_local("__fortran_member_call_result");
        self.emit_u16(Op::LOCAL_SET, result_slot);
        self.emit_u16(Op::LOCAL_GET, receiver_slot);
        if self.compile_assign_target(object).is_err() {
            self.emit(Op::DROP);
        }
        self.emit_u16(Op::LOCAL_GET, result_slot);
    }

    fn try_compile_fortran_derived_type_constructor(
        &mut self,
        name: &str,
        args: &[Argument],
    ) -> Result<bool, String> {
        if self.profile.name != "fortran" || self.has_accessible_local_binding(name) {
            return Ok(false);
        }

        let class_name = self.canon(name);
        let Some(pending) = self.pending_classes.get(&class_name) else {
            return Ok(false);
        };
        if !pending.is_value_type || has_explicit_constructor_signature(self, &class_name) {
            return Ok(false);
        }
        let fields = pending.fields.clone();
        if args.iter().any(|arg| arg.spread) {
            return Ok(false);
        }

        let ctor_global = format!("{}$arity0", class_name);
        if self.defined_globals.contains(&ctor_global) {
            self.emit_var_get(&ctor_global);
        } else {
            self.emit_var_get(name);
        }
        self.emit_u8(Op::CALL_REF, 0);

        let obj_slot = self.define_local("__fortran_type_ctor_obj");
        self.emit_u16(Op::LOCAL_SET, obj_slot);

        let mut positional_index = 0usize;
        for (index, arg) in args.iter().enumerate() {
            let field_name = if let Some(field_name) = arg.name.as_ref() {
                self.canon(field_name)
            } else {
                let Some(field_name) = fields.get(positional_index).cloned() else {
                    return Err(format!(
                        "Fortran type constructor '{}' received too many positional arguments",
                        name
                    ));
                };
                positional_index += 1;
                field_name
            };

            if !fields.iter().any(|field| field == &field_name) {
                return Err(format!(
                    "Fortran type constructor '{}' has no field named '{}'",
                    name, field_name
                ));
            }

            self.compile_expr_with_value_copy(&arg.value)?;
            let value_slot = self.define_local(&format!("__fortran_type_ctor_arg_{}", index));
            self.emit_u16(Op::LOCAL_SET, value_slot);

            let field_idx = self.str_const(&field_name);
            self.emit_u16(Op::LOCAL_GET, obj_slot);
            self.emit_u16(Op::LOCAL_GET, value_slot);
            self.emit_u16(Op::STRUCT_SET, field_idx);
            self.emit(Op::DROP);
        }

        self.emit_u16(Op::LOCAL_GET, obj_slot);
        Ok(true)
    }

    fn overload_type_matches(&self, param_type: &str, arg_type: &str) -> bool {
        let normalized_param =
            Self::normalize_type_hint(strip_generic_suffix(param_type).trim_end_matches('?'));
        let normalized_arg =
            Self::normalize_type_hint(strip_generic_suffix(arg_type).trim_end_matches('?'));
        normalized_param == normalized_arg
            || (Self::is_string_type_hint(&normalized_param)
                && Self::is_string_type_hint(&normalized_arg))
            || (matches!(normalized_param.as_str(), "bool" | "boolean")
                && matches!(normalized_arg.as_str(), "bool" | "boolean"))
            || (is_numeric_overload_type(&normalized_param)
                && is_numeric_overload_type(&normalized_arg))
    }

    pub(super) fn resolve_pending_class_name_for_type_hint(
        &self,
        type_hint: &str,
    ) -> Option<String> {
        let receiver_type = type_hint
            .trim()
            .strip_prefix("class(")
            .and_then(|inner| inner.strip_suffix(')'))
            .map(str::trim)
            .unwrap_or(type_hint.trim());
        let receiver_canon = self.canon(strip_generic_suffix(receiver_type));
        if self.pending_classes.contains_key(&receiver_canon) {
            return Some(receiver_canon);
        }
        self.pending_classes
            .keys()
            .find(|name| {
                name.eq_ignore_ascii_case(receiver_type)
                    || name.eq_ignore_ascii_case(&receiver_canon)
            })
            .cloned()
    }

    fn match_method_overload_chunk(
        &self,
        overloads: &[PendingMethodOverload],
        arg_exprs: &[&Expression],
        include_receiver: bool,
    ) -> Option<usize> {
        self.match_method_overload(overloads, arg_exprs, include_receiver)
            .map(|overload| overload.chunk_idx)
    }

    fn match_method_overload(
        &self,
        overloads: &[PendingMethodOverload],
        arg_exprs: &[&Expression],
        include_receiver: bool,
    ) -> Option<PendingMethodOverload> {
        let effective_args = if include_receiver && !arg_exprs.is_empty() {
            &arg_exprs[1..]
        } else {
            arg_exprs
        };
        let actual_arity = effective_args.len();

        'overload_search: for overload in overloads {
            let signature = &overload.signature;
            let param_count = overload.param_types.len();
            let arity_ok = actual_arity >= signature.min_arity
                && (signature.has_rest || actual_arity <= param_count);
            if !arity_ok {
                continue;
            }

            for (arg_expr, param_type) in effective_args.iter().zip(overload.param_types.iter()) {
                if let Some(arg_type) = self.infer_expr_type_hint(arg_expr) {
                    if !self.overload_type_matches(param_type, &arg_type) {
                        continue 'overload_search;
                    }
                }
            }

            return Some(overload.clone());
        }

        None
    }

    fn resolve_instance_method_overload_chunk(
        &self,
        object: &Expression,
        method_name: &str,
        arg_exprs: &[&Expression],
    ) -> Option<usize> {
        self.resolve_instance_method_overload(object, method_name, arg_exprs, false)
            .map(|overload| overload.chunk_idx)
    }

    fn resolve_instance_method_overload(
        &self,
        object: &Expression,
        method_name: &str,
        arg_exprs: &[&Expression],
        include_receiver: bool,
    ) -> Option<PendingMethodOverload> {
        let receiver_type = resolve_receiver_type_hint(self, object)?;
        let class_name = self.resolve_pending_class_name_for_type_hint(&receiver_type)?;
        let pending = self.pending_classes.get(&class_name)?;
        let method_key = self.js_member_storage_name_for_class(&class_name, method_name);
        let overloads = pending.instance_method_overloads.get(&method_key)?;
        self.match_method_overload(overloads, arg_exprs, include_receiver)
    }

    pub(super) fn pending_class_has_method_name_for_type(
        &self,
        type_hint: &str,
        method_name: &str,
    ) -> bool {
        let Some(class_name) = self.resolve_pending_class_name_for_type_hint(type_hint) else {
            return false;
        };
        let Some(pending) = self.pending_classes.get(&class_name) else {
            return false;
        };
        let method_key = self.js_member_storage_name_for_class(&class_name, method_name);
        pending.static_method_overloads.contains_key(&method_key)
            || pending.instance_method_overloads.contains_key(&method_key)
            || pending
                .static_method_names
                .iter()
                .any(|name| name == &method_key)
            || pending
                .instance_member_names
                .iter()
                .any(|name| name == &method_key)
    }

    fn direct_receiver_has_own_pending_method(
        &self,
        receiver: &Expression,
        method_name: &str,
    ) -> bool {
        let class_name = match &receiver.kind {
            ExprKind::This | ExprKind::Super => self.current_class.clone(),
            ExprKind::Ident(name) => {
                let canon = self.canon(name);
                if canon == self.profile.self_keyword
                    || canon == "me"
                    || canon == "this"
                    || canon == "mybase"
                {
                    self.current_class.clone()
                } else {
                    resolve_receiver_type_hint(self, receiver)
                        .and_then(|hint| self.resolve_pending_class_name_for_type_hint(&hint))
                }
            }
            _ => None,
        };

        let Some(class_name) = class_name else {
            return false;
        };
        let Some(pending) = self.pending_classes.get(&class_name) else {
            return false;
        };

        let method_key = self.js_member_storage_name_for_class(&class_name, method_name);
        pending.instance_method_overloads.contains_key(&method_key)
            || pending
                .instance_member_names
                .iter()
                .any(|name| name == &method_key)
    }

    pub(super) fn resolve_static_method_overload_chunk_for_type(
        &self,
        type_hint: &str,
        method_name: &str,
        arg_exprs: &[&Expression],
    ) -> Option<usize> {
        let class_name = self
            .resolve_pending_class_name_for_type_hint(type_hint)
            .or_else(|| {
                let canon = self.canon(type_hint);
                self.pending_classes.contains_key(&canon).then_some(canon)
            })?;
        let pending = self.pending_classes.get(&class_name)?;
        let method_key = self.js_member_storage_name_for_class(&class_name, method_name);
        let overloads = pending
            .static_method_overloads
            .get(&method_key)
            .or_else(|| pending.instance_method_overloads.get(&method_key))?;
        self.match_method_overload_chunk(overloads, arg_exprs, false)
    }

    fn resolve_static_method_overload_for_type(
        &self,
        type_hint: &str,
        method_name: &str,
        arg_exprs: &[&Expression],
    ) -> Option<PendingMethodOverload> {
        let class_name = self
            .resolve_pending_class_name_for_type_hint(type_hint)
            .or_else(|| {
                let canon = self.canon(type_hint);
                self.pending_classes.contains_key(&canon).then_some(canon)
            })?;
        let pending = self.pending_classes.get(&class_name)?;
        let method_key = self.js_member_storage_name_for_class(&class_name, method_name);
        let overloads = pending
            .static_method_overloads
            .get(&method_key)
            .or_else(|| pending.instance_method_overloads.get(&method_key))?;
        self.match_method_overload(overloads, arg_exprs, false)
    }

    fn resolve_unique_static_method_chunk_for_class(
        &self,
        class_name: &str,
        method_name: &str,
    ) -> Option<usize> {
        let class_name = self.canon(class_name);
        let pending = self.pending_classes.get(&class_name)?;
        let method_key = self.js_member_storage_name_for_class(&class_name, method_name);
        let overloads = pending
            .static_method_overloads
            .get(&method_key)
            .or_else(|| pending.instance_method_overloads.get(&method_key))?;
        (overloads.len() == 1).then_some(overloads[0].chunk_idx)
    }

    fn emit_direct_instance_method_call(
        &mut self,
        chunk_idx: usize,
        method_name: &str,
        obj_tmp: u16,
        args: &[Argument],
        arg_exprs: &[&Expression],
    ) -> Result<(), String> {
        let line = self.line;
        self.emit_u16(Op::REF_FUNC, chunk_idx as u16);
        self.chunk().emit(0, line);
        let fn_tmp = self.define_local("__direct_instance_method_fn");
        self.emit_u16(Op::LOCAL_SET, fn_tmp);

        if let Some(param_modes) = self
            .function_param_modes
            .get(&self.canon(method_name))
            .cloned()
        {
            let receiver_param_offset = usize::from(param_modes.len() == args.len() + 1);
            let user_modes = &param_modes[receiver_param_offset.min(param_modes.len())..];
            if user_modes
                .iter()
                .any(|mode| matches!(mode, PassBy::Ref | PassBy::Out))
            {
                let mut arg_slots = Vec::with_capacity(args.len());
                for (index, arg) in args.iter().enumerate() {
                    self.compile_ref_aware_call_arg(
                        arg,
                        user_modes.get(index).copied().unwrap_or(PassBy::Value),
                    )?;
                    let arg_slot =
                        self.define_local(&format!("__direct_instance_method_arg_{}", index));
                    self.emit_u16(Op::LOCAL_SET, arg_slot);
                    arg_slots.push(arg_slot);
                }

                self.emit_call_ref_with_arg_slots(fn_tmp, Some(obj_tmp), &arg_slots);

                let pack_slot = self.define_local("__direct_instance_method_ref_call_pack");
                self.emit_u16(Op::LOCAL_SET, pack_slot);
                let mut ref_out_index = 1usize;
                for (index, arg) in args.iter().enumerate() {
                    if !matches!(user_modes.get(index), Some(PassBy::Ref | PassBy::Out)) {
                        continue;
                    }
                    self.emit_u16(Op::LOCAL_GET, pack_slot);
                    self.emit_const(Value::F64(ref_out_index as f64));
                    common::collections::emit_get(&mut self.chunks, self.current, self.line);
                    self.compile_assign_target(&arg.value)?;
                    ref_out_index += 1;
                }
                self.emit_u16(Op::LOCAL_GET, pack_slot);
                self.emit_const(Value::F64(0.0));
                common::collections::emit_get(&mut self.chunks, self.current, self.line);
                return Ok(());
            }
        }

        let mut arg_slots = Vec::with_capacity(arg_exprs.len());
        for (index, arg) in arg_exprs.iter().enumerate() {
            self.compile_expr(arg)?;
            let arg_slot = self.define_local(&format!("__direct_instance_method_arg_{}", index));
            self.emit_u16(Op::LOCAL_SET, arg_slot);
            arg_slots.push(arg_slot);
        }
        self.emit_call_ref_with_arg_slots(fn_tmp, Some(obj_tmp), &arg_slots);
        Ok(())
    }

    pub(super) fn emit_direct_static_method_call(
        &mut self,
        chunk_idx: usize,
        arg_exprs: &[&Expression],
    ) -> Result<(), String> {
        let line = self.line;
        self.emit_u16(Op::REF_FUNC, chunk_idx as u16);
        self.chunk().emit(0, line);
        let fn_tmp = self.define_local("__direct_static_method_fn");
        self.emit_u16(Op::LOCAL_SET, fn_tmp);
        let mut arg_slots = Vec::with_capacity(arg_exprs.len());
        for (index, arg) in arg_exprs.iter().enumerate() {
            self.compile_expr(arg)?;
            let arg_slot = self.define_local(&format!("__direct_static_method_arg_{}", index));
            self.emit_u16(Op::LOCAL_SET, arg_slot);
            arg_slots.push(arg_slot);
        }
        self.emit_call_ref_with_arg_slots(fn_tmp, None, &arg_slots);
        Ok(())
    }

    pub(super) fn select_call_signature<'a>(
        &self,
        signatures: &'a [CallSignature],
        args: &[Argument],
    ) -> Option<&'a CallSignature> {
        let mut rest_candidate = None;
        for signature in signatures {
            let fits = args.len() >= signature.min_arity
                && (signature.has_rest || args.len() <= signature.param_names.len());
            if !fits {
                continue;
            }
            if signature.has_rest {
                rest_candidate.get_or_insert(signature);
            } else {
                return Some(signature);
            }
        }
        rest_candidate
    }

    pub(super) fn emit_stamp_rest_metadata_on_stack(&mut self, fixed_count: usize) {
        let key = self.str_const("__vybe_rest_fixed_arity");
        inst!(self, core_wasm::dup);
        self.emit_const(Value::F64(fixed_count as f64));
        self.emit_u16(Op::STRUCT_SET, key);
        self.emit(Op::DROP);
    }

    fn bind_js_this_for_call(
        &mut self,
        receiver_slot: Option<u16>,
        saved_name: &str,
    ) -> Option<u16> {
        let saved_js_this = self.save_js_this(saved_name);
        if !self.is_js_profile() {
            return saved_js_this;
        }

        if let Some(slot) = receiver_slot {
            self.emit_u16(Op::LOCAL_GET, slot);
            self.set_js_this_from_stack();
            return saved_js_this;
        }

        let js_global_this = self.str_const("__js_global_this");
        self.emit_u16(Op::GLOBAL_GET, js_global_this);
        let global_this_slot = self.define_local("__js_global_this_value");
        self.emit_u16(Op::LOCAL_SET, global_this_slot);
        self.emit_u16(Op::LOCAL_GET, global_this_slot);
        self.emit(Op::REF_IS_NULL);
        let line = self.line;
        self.chunk().emit_if_value(line);
        self.emit_common("object.new", 0, line);
        inst!(self, core_wasm::dup);
        self.emit_u16(Op::GLOBAL_SET, js_global_this);
        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, global_this_slot);
        fn_call!(self, "wasm:js-undefined", "test", 1);
        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if_value(line);
        self.emit_common("object.new", 0, self.line);
        inst!(self, core_wasm::dup);
        self.emit_u16(Op::GLOBAL_SET, js_global_this);
        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, global_this_slot);
        self.chunk().emit_end(line);
        self.chunk().emit_end(line);
        self.set_js_this_from_stack();
        saved_js_this
    }

    fn restore_js_this_after_call(&mut self, saved_js_this: Option<u16>, result_local_name: &str) {
        let Some(_) = saved_js_this else {
            return;
        };
        let result_slot = self.define_local(result_local_name);
        self.emit_u16(Op::LOCAL_SET, result_slot);
        self.restore_js_this(saved_js_this);
        self.emit_u16(Op::LOCAL_GET, result_slot);
    }

    fn emit_normal_call_from_arg_slots(
        &mut self,
        callee_slot: u16,
        receiver_slot: Option<u16>,
        js_this_slot: Option<u16>,
        arg_slots: &[u16],
    ) {
        let saved_js_this =
            self.bind_js_this_for_call(js_this_slot.or(receiver_slot), "__js_prev_this_arg_call");
        self.emit_u16(Op::LOCAL_GET, callee_slot);
        if let Some(receiver_slot) = receiver_slot {
            self.emit_u16(Op::LOCAL_GET, receiver_slot);
        }
        for slot in arg_slots {
            self.emit_u16(Op::LOCAL_GET, *slot);
        }
        if self.profile.name == "fortran" && receiver_slot.is_none() && arg_slots.len() == 1 {
            inst!(self, core_wasm::undefined);
            self.emit_u8(Op::CALL_REF, 2);
        } else {
            self.emit_u8(
                Op::CALL_REF,
                (arg_slots.len() + usize::from(receiver_slot.is_some())) as u8,
            );
        }
        self.restore_js_this_after_call(saved_js_this, "__js_arg_call_result");
    }

    fn emit_rest_call_from_arg_slots(
        &mut self,
        callee_slot: u16,
        receiver_slot: Option<u16>,
        js_this_slot: Option<u16>,
        arg_slots: &[u16],
        fixed_count: usize,
    ) {
        let saved_js_this = self.bind_js_this_for_call(
            js_this_slot.or(receiver_slot),
            "__js_prev_this_rest_arg_call",
        );
        let argc = fixed_count + 1 + usize::from(receiver_slot.is_some());
        let line = self.line;
        self.emit_u16(Op::LOCAL_GET, callee_slot);
        if let Some(receiver_slot) = receiver_slot {
            self.emit_u16(Op::LOCAL_GET, receiver_slot);
        }
        for index in 0..fixed_count {
            if let Some(slot) = arg_slots.get(index) {
                self.emit_u16(Op::LOCAL_GET, *slot);
            } else {
                inst!(self, core_wasm::undefined);
            }
        }
        let rest_slot = self.define_local("__runtime_rest_call_array");
        common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
        self.emit_u16(Op::LOCAL_SET, rest_slot);
        for slot in arg_slots.iter().skip(fixed_count) {
            self.emit_u16(Op::LOCAL_GET, rest_slot);
            self.emit_u16(Op::LOCAL_GET, *slot);
            common::collections::emit_push(&mut self.chunks, self.current, line);
            self.emit(Op::DROP);
        }
        self.emit_u16(Op::LOCAL_GET, rest_slot);
        self.emit_u8(Op::CALL_REF, argc as u8);
        self.restore_js_this_after_call(saved_js_this, "__js_rest_arg_call_result");
    }

    pub(super) fn emit_array_value_or_undefined(
        &mut self,
        args_slot: u16,
        len_slot: u16,
        index: usize,
    ) {
        self.emit_u16(Op::LOCAL_GET, len_slot);
        self.emit_const(Value::F64(index as f64));
        {
            let line = self.line;
            crate::emitter::ops::emit_dyn_gt(self.chunk(), line);
        };
        let line = self.line;
        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if_value(line);
        self.emit_u16(Op::LOCAL_GET, args_slot);
        self.emit_const(Value::F64(index as f64));
        common::collections::emit_get(&mut self.chunks, self.current, self.line);
        self.chunk().emit_else(line);
        inst!(self, core_wasm::undefined);
        self.chunk().emit_end(line);
    }

    fn compile_out_call_arg(&mut self, arg: &Argument) -> Result<(), String> {
        if self.profile.name == "fortran"
            && (self.expr_is_array_like(&arg.value)
                || self
                    .infer_expr_type_hint(&arg.value)
                    .as_deref()
                    .and_then(Self::fortran_out_param_ctor_name)
                    .is_some())
        {
            self.compile_expr_with_value_copy(&arg.value)?;
        } else {
            self.emit(Op::NULL);
        }
        Ok(())
    }

    fn compile_ref_aware_call_arg(&mut self, arg: &Argument, mode: PassBy) -> Result<(), String> {
        match mode {
            PassBy::Out => self.compile_out_call_arg(arg)?,
            PassBy::Ref | PassBy::Const if self.profile.name == "fortran" => {
                self.compile_expr(&arg.value)?;
            }
            PassBy::Ref | PassBy::Const | PassBy::Value => {
                self.compile_expr_with_value_copy(&arg.value)?
            }
        }
        Ok(())
    }

    fn mode_needs_ref_aware_call_handling(&self, mode: PassBy) -> bool {
        matches!(mode, PassBy::Ref | PassBy::Out)
            || (self.profile.name == "fortran" && matches!(mode, PassBy::Const))
    }

    fn mode_needs_call_writeback(&self, mode: PassBy) -> bool {
        matches!(mode, PassBy::Ref | PassBy::Out)
    }

    fn emit_normal_call_from_args_array(
        &mut self,
        callee_slot: u16,
        receiver_slot: Option<u16>,
        args_slot: u16,
        _known_len: Option<usize>,
    ) {
        let saved_js_this = self.bind_js_this_for_call(receiver_slot, "__js_prev_this_array_call");
        self.emit_u16(Op::LOCAL_GET, callee_slot);
        if let Some(rs) = receiver_slot {
            self.emit_u16(Op::LOCAL_GET, rs);
        } else {
            inst!(self, core_wasm::undefined);
        }
        self.emit_u16(Op::LOCAL_GET, args_slot);
        fn_call!(self, "ecma:function", "apply", 3);
        self.restore_js_this_after_call(saved_js_this, "__js_array_call_result");
    }

    fn emit_rest_call_from_args_array(
        &mut self,
        callee_slot: u16,
        receiver_slot: Option<u16>,
        args_slot: u16,
        known_len: Option<usize>,
        fixed_count: usize,
    ) {
        let argc = fixed_count + 1 + usize::from(receiver_slot.is_some());
        let line = self.line;
        self.emit_u16(Op::LOCAL_GET, callee_slot);
        if let Some(receiver_slot) = receiver_slot {
            self.emit_u16(Op::LOCAL_GET, receiver_slot);
        }

        match known_len {
            Some(known_len) => {
                for index in 0..fixed_count {
                    if index < known_len {
                        self.emit_u16(Op::LOCAL_GET, args_slot);
                        self.emit_const(Value::F64(index as f64));
                        common::collections::emit_get(&mut self.chunks, self.current, line);
                    } else {
                        inst!(self, core_wasm::undefined);
                    }
                }
                self.emit_u16(Op::LOCAL_GET, args_slot);
                self.emit_const(Value::F64(fixed_count as f64));
                self.emit_const(Value::F64(known_len as f64));
                common::collections::emit_slice(&mut self.chunks, self.current, line);
            }
            None => {
                let len_slot = self.define_local("__runtime_spread_len");
                self.emit_u16(Op::LOCAL_GET, args_slot);
                common::collections::emit_len(&mut self.chunks, self.current, line);
                self.emit_u16(Op::LOCAL_SET, len_slot);
                for index in 0..fixed_count {
                    self.emit_array_value_or_undefined(args_slot, len_slot, index);
                }
                self.emit_u16(Op::LOCAL_GET, args_slot);
                self.emit_const(Value::F64(fixed_count as f64));
                self.emit_u16(Op::LOCAL_GET, len_slot);
                common::collections::emit_slice(&mut self.chunks, self.current, line);
            }
        }

        self.emit_u8(Op::CALL_REF, argc as u8);
    }

    fn emit_dispatch_and_store_from_arg_slots(
        &mut self,
        callee_slot: u16,
        receiver_slot: Option<u16>,
        js_this_slot: Option<u16>,
        arg_slots: &[u16],
        result_slot: u16,
    ) {
        // Proxy modules: the callee may be a Proxy whose apply trap must
        // fire (ECMA-262 §10.5.12). ecma:proxy.apply falls through to an
        // ordinary invoke for plain callables, so all dynamic calls can
        // route through it.
        if self.is_js_profile() && self.uses_proxy && receiver_slot.is_none() {
            let line = self.line;
            let args_arr_slot = self.define_local("__proxy_apply_args");
            common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
            self.emit_u16(Op::LOCAL_SET, args_arr_slot);
            for slot in arg_slots {
                self.emit_u16(Op::LOCAL_GET, args_arr_slot);
                self.emit_u16(Op::LOCAL_GET, *slot);
                common::collections::emit_push(&mut self.chunks, self.current, line);
                self.emit(Op::DROP);
            }
            self.emit_u16(Op::LOCAL_GET, callee_slot);
            if let Some(this_slot) = js_this_slot {
                self.emit_u16(Op::LOCAL_GET, this_slot);
            } else {
                inst!(self, core_wasm::undefined);
            }
            self.emit_u16(Op::LOCAL_GET, args_arr_slot);
            let idx = self.import("ecma:proxy", "apply");
            self.emit_host_call(idx, 3);
            self.emit_u16(Op::LOCAL_SET, result_slot);
            return;
        }

        let rest_fixed_counts: Vec<u8> = self.rest_fixed_arities.iter().copied().collect();
        if rest_fixed_counts.is_empty() {
            self.emit_normal_call_from_arg_slots(
                callee_slot,
                receiver_slot,
                js_this_slot,
                arg_slots,
            );
            self.emit_u16(Op::LOCAL_SET, result_slot);
            return;
        }

        let rest_key = self.str_const("__vybe_rest_fixed_arity");
        let rest_arity_slot = self.define_local("__call_rest_fixed_arity");
        self.emit_u16(Op::LOCAL_GET, callee_slot);
        self.emit_u16(Op::STRUCT_GET, rest_key);
        self.emit_u16(Op::LOCAL_SET, rest_arity_slot);

        let used_rest_slot = self.define_local("__call_used_rest_arity");
        self.emit_const(Value::I32(0));
        self.emit_u16(Op::LOCAL_SET, used_rest_slot);

        self.emit_u16(Op::LOCAL_GET, rest_arity_slot);
        self.emit(Op::REF_IS_NULL);
        let line = self.line;
        self.chunk().emit_if(line);
        self.chunk().emit_else(line);
        for fixed_count in rest_fixed_counts {
            self.emit_u16(Op::LOCAL_GET, used_rest_slot);
            self.emit(Op::I32_EQZ);
            let line = self.line;
            self.chunk().emit_if(line);
            self.emit_u16(Op::LOCAL_GET, rest_arity_slot);
            self.emit_const(Value::F64(fixed_count as f64));
            {
                let line = self.line;
                crate::emitter::ops::emit_dyn_eq(self.chunk(), line);
            };
            let line = self.line;
            crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
            self.chunk().emit_if(line);
            self.emit_rest_call_from_arg_slots(
                callee_slot,
                receiver_slot,
                js_this_slot,
                arg_slots,
                fixed_count as usize,
            );
            self.emit_u16(Op::LOCAL_SET, result_slot);
            self.emit_const(Value::I32(1));
            self.emit_u16(Op::LOCAL_SET, used_rest_slot);
            self.chunk().emit_end(line);
            self.chunk().emit_end(line);
        }
        self.chunk().emit_end(line);

        self.emit_u16(Op::LOCAL_GET, used_rest_slot);
        self.emit(Op::I32_EQZ);
        let line = self.line;
        self.chunk().emit_if(line);
        self.emit_normal_call_from_arg_slots(callee_slot, receiver_slot, js_this_slot, arg_slots);
        self.emit_u16(Op::LOCAL_SET, result_slot);
        self.chunk().emit_end(line);
    }

    fn emit_call_ref_with_arg_slots(
        &mut self,
        callee_slot: u16,
        receiver_slot: Option<u16>,
        arg_slots: &[u16],
    ) {
        let result_slot = self.define_local("__call_runtime_result");
        if let Some(receiver_slot) = receiver_slot {
            self.emit_u16(Op::LOCAL_GET, receiver_slot);
            self.emit(Op::REF_IS_NULL);
            let line = self.line;
            self.chunk().emit_if(line);
            self.emit_dispatch_and_store_from_arg_slots(
                callee_slot,
                None,
                None,
                arg_slots,
                result_slot,
            );
            self.chunk().emit_else(line);
            self.emit_u16(Op::LOCAL_GET, receiver_slot);
            fn_call!(self, "wasm:js-undefined", "test", 1);
            let line = self.line;
            self.chunk().emit_if(line);
            self.emit_dispatch_and_store_from_arg_slots(
                callee_slot,
                None,
                None,
                arg_slots,
                result_slot,
            );
            self.chunk().emit_else(line);
            self.emit_dispatch_and_store_from_arg_slots(
                callee_slot,
                Some(receiver_slot),
                Some(receiver_slot),
                arg_slots,
                result_slot,
            );
            self.chunk().emit_end(line);
            self.chunk().emit_end(line);
        } else {
            self.emit_dispatch_and_store_from_arg_slots(
                callee_slot,
                None,
                None,
                arg_slots,
                result_slot,
            );
        }
        self.emit_u16(Op::LOCAL_GET, result_slot);
    }

    fn finish_member_index_call_path(
        &mut self,
        callee: &Expression,
        arg_exprs: &[&Expression],
        fn_tmp: u16,
        line: u32,
    ) -> Result<(), String> {
        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, fn_tmp);
        for arg in arg_exprs {
            self.compile_array_index_operand_for_owner(callee, arg)?;
            let line = self.line;
            common::collections::emit_get(&mut self.chunks, self.current, line);
        }
        self.chunk().emit_end(line);
        Ok(())
    }

    fn emit_call_ref_with_bound_js_this_arg_slots(
        &mut self,
        callee_slot: u16,
        js_this_slot: u16,
        arg_slots: &[u16],
    ) {
        let result_slot = self.define_local("__call_runtime_result");
        self.emit_dispatch_and_store_from_arg_slots(
            callee_slot,
            None,
            Some(js_this_slot),
            arg_slots,
            result_slot,
        );
        self.emit_u16(Op::LOCAL_GET, result_slot);
    }

    fn emit_js_lookup_or_invoke_method_call(
        &mut self,
        obj_slot: u16,
        method_name: &str,
        arg_slots: &[u16],
    ) {
        let lookup = self.import("ecma:value", "getMethodForCall");
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_const(Value::String(Arc::from(method_name)));
        self.emit_host_call(lookup, 2);
        let lookup_slot = self.define_local("__js_lookup_fn");
        self.emit_u16(Op::LOCAL_SET, lookup_slot);

        self.emit_u16(Op::LOCAL_GET, lookup_slot);
        self.emit(Op::REF_IS_NULL);
        let line = self.line;
        self.chunk().emit_if_value(line);
        self.emit_js_invoke_method_call(obj_slot, method_name, arg_slots);
        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, lookup_slot);
        fn_call!(self, "wasm:js-undefined", "test", 1);
        let line = self.line;
        self.chunk().emit_if_value(line);
        self.emit_js_invoke_method_call(obj_slot, method_name, arg_slots);
        self.chunk().emit_else(line);
        self.emit_call_ref_with_bound_js_this_arg_slots(lookup_slot, obj_slot, arg_slots);
        self.chunk().emit_end(line);
        self.chunk().emit_end(line);
    }

    fn emit_js_invoke_method_call(&mut self, obj_slot: u16, method_name: &str, arg_slots: &[u16]) {
        let invoke = self.import("ecma:value", "invokeMethod");
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_const(Value::String(Arc::from(method_name)));
        for slot in arg_slots {
            self.emit_u16(Op::LOCAL_GET, *slot);
        }
        self.emit_host_call(invoke, (arg_slots.len() + 2) as u8);
    }

    fn emit_js_invoke_method_from_args_array(
        &mut self,
        obj_slot: u16,
        method_name: &str,
        args_slot: u16,
        _known_len: Option<usize>,
    ) {
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        inst!(self, core_wasm::string_const, method_name);
        fn_call!(self, "ecma:value", "getMethodForCall", 2);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_u16(Op::LOCAL_GET, args_slot);
        fn_call!(self, "ecma:function", "apply", 3);
    }

    fn emit_dispatch_and_store_from_args_array(
        &mut self,
        callee_slot: u16,
        receiver_slot: Option<u16>,
        args_slot: u16,
        known_len: Option<usize>,
        result_slot: u16,
    ) {
        let rest_fixed_counts: Vec<u8> = self.rest_fixed_arities.iter().copied().collect();
        if !rest_fixed_counts.is_empty() {
            let rest_key = self.str_const("__vybe_rest_fixed_arity");
            let rest_arity_slot = self.define_local("__spread_rest_fixed_arity");
            self.emit_u16(Op::LOCAL_GET, callee_slot);
            self.emit_u16(Op::STRUCT_GET, rest_key);
            self.emit_u16(Op::LOCAL_SET, rest_arity_slot);

            let used_rest_slot = self.define_local("__spread_used_rest_arity");
            self.emit_const(Value::I32(0));
            self.emit_u16(Op::LOCAL_SET, used_rest_slot);

            self.emit_u16(Op::LOCAL_GET, rest_arity_slot);
            self.emit(Op::REF_IS_NULL);
            let line = self.line;
            self.chunk().emit_if(line);
            self.chunk().emit_else(line);
            for fixed_count in rest_fixed_counts {
                self.emit_u16(Op::LOCAL_GET, used_rest_slot);
                self.emit(Op::I32_EQZ);
                let line = self.line;
                self.chunk().emit_if(line);
                self.emit_u16(Op::LOCAL_GET, rest_arity_slot);
                self.emit_const(Value::F64(fixed_count as f64));
                {
                    let line = self.line;
                    crate::emitter::ops::emit_dyn_eq(self.chunk(), line);
                };
                let line = self.line;
                crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                self.chunk().emit_if(line);
                self.emit_rest_call_from_args_array(
                    callee_slot,
                    receiver_slot,
                    args_slot,
                    known_len,
                    fixed_count as usize,
                );
                self.emit_u16(Op::LOCAL_SET, result_slot);
                self.emit_const(Value::I32(1));
                self.emit_u16(Op::LOCAL_SET, used_rest_slot);
                self.chunk().emit_end(line);
                self.chunk().emit_end(line);
            }
            self.chunk().emit_end(line);

            self.emit_u16(Op::LOCAL_GET, used_rest_slot);
            self.emit(Op::I32_EQZ);
            let line = self.line;
            self.chunk().emit_if(line);
            self.emit_normal_call_from_args_array(callee_slot, receiver_slot, args_slot, known_len);
            self.emit_u16(Op::LOCAL_SET, result_slot);
            self.chunk().emit_end(line);
            return;
        }

        self.emit_normal_call_from_args_array(callee_slot, receiver_slot, args_slot, known_len);
        self.emit_u16(Op::LOCAL_SET, result_slot);
    }

    fn emit_call_ref_with_args_array(
        &mut self,
        callee_slot: u16,
        receiver_slot: Option<u16>,
        args_slot: u16,
        known_len: Option<usize>,
    ) {
        let result_slot = self.define_local("__spread_call_runtime_result");
        if let Some(receiver_slot) = receiver_slot {
            self.emit_u16(Op::LOCAL_GET, receiver_slot);
            self.emit(Op::REF_IS_NULL);
            let line = self.line;
            self.chunk().emit_if(line);
            self.emit_dispatch_and_store_from_args_array(
                callee_slot,
                None,
                args_slot,
                known_len,
                result_slot,
            );
            self.chunk().emit_else(line);
            self.emit_u16(Op::LOCAL_GET, receiver_slot);
            fn_call!(self, "wasm:js-undefined", "test", 1);
            let line = self.line;
            self.chunk().emit_if(line);
            self.emit_dispatch_and_store_from_args_array(
                callee_slot,
                None,
                args_slot,
                known_len,
                result_slot,
            );
            self.chunk().emit_else(line);
            self.emit_dispatch_and_store_from_args_array(
                callee_slot,
                Some(receiver_slot),
                args_slot,
                known_len,
                result_slot,
            );
            self.chunk().emit_end(line);
            self.chunk().emit_end(line);
        } else {
            self.emit_dispatch_and_store_from_args_array(
                callee_slot,
                None,
                args_slot,
                known_len,
                result_slot,
            );
        }
        self.emit_u16(Op::LOCAL_GET, result_slot);
    }

    /// Spread argument to a `host:*`-bound builtin (`String.raw({raw}, ...v)`).
    ///
    /// A fixed-argc CALL_IMPORT cannot represent a runtime-length argument
    /// list, so route through ECMA's own variadic call primitive instead:
    /// §13.3.8.1 ArgumentListEvaluation collects the full list into an array
    /// (spread-aware concat — `compile_call_args_array`), then §28.1.1
    /// `Reflect.apply(target, undefined, argsList)` makes the call, and the
    /// spec-shaped host fn receives individually-expanded arguments.
    fn try_compile_spread_host_builtin(
        &mut self,
        callee: &Expression,
        name: &str,
        args: &[Argument],
    ) -> Result<bool, String> {
        if !self.is_js_profile() || !args.iter().any(|a| a.spread) {
            return Ok(false);
        }
        let Some(def) = self.profile.lookup_builtin(name) else {
            return Ok(false);
        };
        if !matches!(&def.emit, BuiltinEmit::HostCall(..)) {
            return Ok(false);
        }
        // §13.3.6: evaluate the callee reference before the arguments.
        self.compile_expr(callee)?;
        let callee_slot = self.define_local("__spread_builtin_callee");
        self.emit_u16(Op::LOCAL_SET, callee_slot);
        let (args_slot, _) = self.compile_call_args_array(args, "spread_builtin")?;
        self.emit_u16(Op::LOCAL_GET, callee_slot);
        inst!(self, core_wasm::undefined);
        self.emit_u16(Op::LOCAL_GET, args_slot);
        let idx = self.import("ecma:reflect", "apply");
        self.emit_host_call(idx, 3);
        Ok(true)
    }

    pub(super) fn compile_call_args_array(
        &mut self,
        args: &[Argument],
        local_prefix: &str,
    ) -> Result<(u16, Option<usize>), String> {
        let line = self.line;
        let args_slot = self.define_local(&format!("__{}_args", local_prefix));
        common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
        self.emit_u16(Op::LOCAL_SET, args_slot);

        let mut known_len: Option<usize> = Some(0);
        for arg in args {
            if arg.spread {
                self.emit_u16(Op::LOCAL_GET, args_slot);
                self.compile_expr(&arg.value)?;
                common::collections::emit_concat(&mut self.chunks, self.current, line);
                self.emit_u16(Op::LOCAL_SET, args_slot);
                if let ExprKind::Array(elements) = &arg.value.kind {
                    if let Some(len) = known_len.as_mut() {
                        *len += elements.len();
                    }
                } else {
                    known_len = None;
                }
            } else {
                self.emit_u16(Op::LOCAL_GET, args_slot);
                self.compile_expr(&arg.value)?;
                common::collections::emit_push(&mut self.chunks, self.current, line);
                self.emit(Op::DROP);
                if let Some(len) = known_len.as_mut() {
                    *len += 1;
                }
            }
        }

        Ok((args_slot, known_len))
    }

    fn expr_is_known_js_promise_like(&self, expr: &Expression) -> bool {
        if self
            .infer_expr_type_hint(expr)
            .as_deref()
            .map(Compiler::normalize_type_hint)
            .is_some_and(|type_hint: String| type_hint.eq_ignore_ascii_case("promise"))
        {
            return true;
        }

        match &expr.kind {
            ExprKind::New { class, .. } => {
                matches!(&class.kind, ExprKind::Ident(name) if name == "Promise")
            }
            ExprKind::Call { callee, .. } => match &callee.kind {
                ExprKind::Member { object, field, .. } => {
                    if matches!(&object.kind, ExprKind::Ident(name) if name == "Promise") {
                        matches!(
                            field.as_str(),
                            "resolve"
                                | "reject"
                                | "all"
                                | "race"
                                | "allSettled"
                                | "any"
                                | "try"
                                | "withResolvers"
                        )
                    } else if matches!(&object.kind, ExprKind::Ident(name) if name == "Array") {
                        field == "fromAsync"
                    } else {
                        matches!(field.as_str(), "then" | "catch" | "finally")
                            && self.expr_is_known_js_promise_like(object)
                    }
                }
                ExprKind::Ident(name) => self
                    .function_return_types
                    .get(&self.canon(name))
                    .is_some_and(|ty| {
                        Compiler::normalize_type_hint(ty).eq_ignore_ascii_case("promise")
                    }),
                _ => false,
            },
            _ => false,
        }
    }

    #[allow(dead_code)]
    fn compile_js_promise_chain_wrapper(
        &mut self,
        kind: JsPromiseChainKind,
    ) -> Result<usize, String> {
        let params = js_promise_chain_params(kind);
        let func_idx = self.chunks.len();
        let mut chunk =
            common::functions::create_function_chunk("<js_promise_chain>", params.len() as u8);
        chunk.is_async = true;
        self.chunks.push(chunk);
        self.scopes.push(Scope::new_function());

        let saved_current = self.current;
        self.current = func_idx;
        let saved_label_base = self.function_label_base;
        self.function_label_base = self.label_depth;
        let saved_func_name = self.current_func_name.take();
        self.current_func_name = Some("<js_promise_chain>".into());
        let saved_result_slot = self.current_result_slot.take();
        let saved_ref_out = self.current_ref_out_params.take();
        self.current_ref_out_params = None;

        for param in params {
            self.define_local(param);
        }

        let async_try = {
            let line = self.line;
            common::functions::emit_async_body_start(&mut self.chunks[self.current], line)
        };
        self.active_async_try_depth += 1;
        let body = js_promise_chain_body(kind);
        for statement in &body {
            self.compile_stmt(statement)?;
        }
        self.active_async_try_depth = self.active_async_try_depth.saturating_sub(1);

        let line = self.line;
        common::functions::emit_async_body_fallthrough(
            &mut self.chunks[self.current],
            async_try,
            line,
        );
        self.emit_return();
        common::functions::patch_async_body_catch(&mut self.chunks[self.current], async_try);
        self.emit_return();

        let ns = self.scope().next_slot;
        self.chunks[func_idx].finalize_local_count(ns);
        self.scopes.pop();
        self.current = saved_current;
        self.function_label_base = saved_label_base;
        self.current_func_name = saved_func_name;
        self.current_result_slot = saved_result_slot;
        self.current_ref_out_params = saved_ref_out;
        Ok(func_idx)
    }

    fn try_compile_js_promise_chain_call(
        &mut self,
        object: &Expression,
        field: &str,
        arg_exprs: &[&Expression],
    ) -> Result<bool, String> {
        // Promise.prototype.then/catch/finally route to the ECMA host promise
        // engine (`ecma:promise.{then,catch,finally}` → then_impl/finally_impl),
        // which registers reactions and schedules them as microtasks per
        // ECMA-262 §27.2 (settled → enqueue, pending → reaction list), and
        // shares the same settle path that resumes JSPI-awaiting fibers. This is
        // the single spec-correct engine; the old eager JSPI-await wrappers ran
        // reactions synchronously at the call site (wrong ordering) and are gone.
        let is_promise_like = self.expr_is_known_js_promise_like(object);
        if !is_promise_like {
            return Ok(false);
        }

        // Callback params the host fn reads after the receiver promise:
        //   then(promise, onFulfilled, onRejected) — 2
        //   catch(promise, onRejected)             — 1
        //   finally(promise, onFinally)            — 1
        let callback_count: usize = match field {
            "then" => 2,
            "catch" | "finally" => 1,
            _ => return Ok(false),
        };

        self.compile_expr(object)?; // receiver promise = args[0]
        for arg in arg_exprs.iter().take(callback_count) {
            self.compile_expr(arg)?;
        }
        for _ in arg_exprs.len().min(callback_count)..callback_count {
            self.emit(Op::NULL); // pad omitted callbacks (e.g. `.then(onF)`)
        }
        let idx = self.import("ecma:promise", field);
        self.emit_host_call(idx, (1 + callback_count) as u8);
        Ok(true)
    }

    fn emit_flat_call_args_array(
        &mut self,
        args: &[Argument],
        slot_name: &str,
    ) -> Result<u16, String> {
        let line = self.line;
        let args_slot = self.define_local(slot_name);
        common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
        self.emit_u16(Op::LOCAL_SET, args_slot);
        for arg in args {
            if arg.spread {
                self.emit_u16(Op::LOCAL_GET, args_slot);
                self.compile_expr(&arg.value)?;
                common::collections::emit_concat(&mut self.chunks, self.current, line);
                self.emit_u16(Op::LOCAL_SET, args_slot);
            } else {
                self.emit_u16(Op::LOCAL_GET, args_slot);
                self.compile_expr_with_value_copy(&arg.value)?;
                common::collections::emit_push(&mut self.chunks, self.current, line);
                self.emit(Op::DROP);
            }
        }
        Ok(args_slot)
    }

    pub(super) fn emit_known_rest_call_from_local(
        &mut self,
        callee_slot: u16,
        receiver_slot: Option<u16>,
        args: &[Argument],
        signature: &CallSignature,
    ) -> Result<(), String> {
        let fixed_count = signature.param_names.len().saturating_sub(1);
        let argc = fixed_count + 1 + usize::from(receiver_slot.is_some());

        self.emit_u16(Op::LOCAL_GET, callee_slot);
        if let Some(receiver_slot) = receiver_slot {
            self.emit_u16(Op::LOCAL_GET, receiver_slot);
        }

        if fixed_count == 0 && args.len() == 1 && args[0].spread {
            self.compile_expr(&args[0].value)?;
            self.emit_u8(Op::CALL_REF, argc as u8);
            return Ok(());
        }

        if args.iter().any(|arg| arg.spread) {
            let line = self.line;
            let args_slot = self.emit_flat_call_args_array(args, "__packed_rest_call_args")?;
            for index in 0..fixed_count {
                self.emit_u16(Op::LOCAL_GET, args_slot);
                self.emit_const(Value::F64(index as f64));
                common::collections::emit_get(&mut self.chunks, self.current, line);
            }
            self.emit_u16(Op::LOCAL_GET, args_slot);
            self.emit_const(Value::F64(fixed_count as f64));
            self.emit_u16(Op::LOCAL_GET, args_slot);
            common::collections::emit_len(&mut self.chunks, self.current, line);
            common::collections::emit_slice(&mut self.chunks, self.current, line);
        } else {
            for index in 0..fixed_count {
                if let Some(arg) = args.get(index) {
                    self.compile_expr_with_value_copy(&arg.value)?;
                } else {
                    inst!(self, core_wasm::undefined);
                }
            }

            let line = self.line;
            let rest_slot = self.define_local("__packed_rest_call_array");
            common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
            self.emit_u16(Op::LOCAL_SET, rest_slot);
            for arg in args.iter().skip(fixed_count) {
                self.emit_u16(Op::LOCAL_GET, rest_slot);
                self.compile_expr_with_value_copy(&arg.value)?;
                common::collections::emit_push(&mut self.chunks, self.current, line);
                self.emit(Op::DROP);
            }
            self.emit_u16(Op::LOCAL_GET, rest_slot);
        }

        self.emit_u8(Op::CALL_REF, argc as u8);
        Ok(())
    }

    fn emit_variadic_array_call_from_local(
        &mut self,
        callee_slot: u16,
        array_expr: &Expression,
    ) -> Result<(), String> {
        self.compile_expr(array_expr)?;
        let args_slot = self.define_local("__params_array_args");
        self.emit_u16(Op::LOCAL_SET, args_slot);

        self.emit_call_ref_with_args_array(callee_slot, None, args_slot, None);
        Ok(())
    }

    pub(super) fn js_error_instanceof_chain(type_name: &str) -> &'static [&'static str] {
        match type_name.trim() {
            "Error" => &["Error"],
            "EvalError" => &["EvalError", "Error"],
            "RangeError" => &["RangeError", "Error"],
            "ReferenceError" => &["ReferenceError", "Error"],
            "SyntaxError" => &["SyntaxError", "Error"],
            "TypeError" => &["TypeError", "Error"],
            "URIError" => &["URIError", "Error"],
            "AggregateError" => &["AggregateError", "Error"],
            _ => &[],
        }
    }

    pub(super) fn emit_js_exception_ctor_from_message_value(
        &mut self,
        type_name: &str,
    ) -> Result<(), String> {
        let msg_val = self.define_local("__exc_msg_val");
        self.emit_u16(Op::LOCAL_SET, msg_val);

        self.emit_u16(Op::STRUCT_NEW, 0);
        inst!(self, core_wasm::dup);
        self.emit_u16(Op::LOCAL_GET, msg_val);
        let line = self.line;
        common::errors::emit_exception_new_finalize(self.chunk(), type_name, line);

        let exc_tmp = self.define_local("__exc_tmp");
        self.emit_u16(Op::LOCAL_SET, exc_tmp);

        if self.is_js_profile() {
            // Fix property descriptors to be non-enumerable per ECMA-262 §20.5.
            // message, name, and internal properties (__type, __exception_type) should be non-enumerable.
            let define_prop_idx = self.import("ecma:object", "defineProperty");

            for prop_name in &["__type", "__exception_type", "message", "name"] {
                self.emit_u16(Op::LOCAL_GET, exc_tmp);
                inst!(self, core_wasm::string_const, prop_name);
                common::dict::emit_new(&mut self.chunks, self.current, line);
                inst!(self, core_wasm::dup);
                self.emit_u16(Op::LOCAL_GET, exc_tmp);
                let prop_key = self.str_const(prop_name);
                self.emit_u16(Op::STRUCT_GET, prop_key);
                let val_key = self.str_const("value");
                self.emit_u16(Op::STRUCT_SET, val_key);
                self.emit(Op::DROP);
                inst!(self, core_wasm::dup);
                self.emit_const(Value::Bool(false));
                let enum_key = self.str_const("enumerable");
                self.emit_u16(Op::STRUCT_SET, enum_key);
                self.emit(Op::DROP);
                self.emit_host_call(define_prop_idx, 3);
                self.emit(Op::DROP);
            }

            // Set tostringtag as non-enumerable
            self.emit_u16(Op::LOCAL_GET, exc_tmp);
            inst!(self, core_wasm::string_const, "tostringtag");
            common::dict::emit_new(&mut self.chunks, self.current, line);
            inst!(self, core_wasm::dup);
            inst!(self, core_wasm::string_const, "Error");
            let val_key = self.str_const("value");
            self.emit_u16(Op::STRUCT_SET, val_key);
            self.emit(Op::DROP);
            inst!(self, core_wasm::dup);
            inst!(self, core_wasm::bool_const, false);
            let enum_key = self.str_const("enumerable");
            self.emit_u16(Op::STRUCT_SET, enum_key);
            self.emit(Op::DROP);
            self.emit_host_call(define_prop_idx, 3);
            self.emit(Op::DROP);
        }

        self.emit_const(Value::String(Arc::from(format!("{}: ", type_name))));
        self.emit_u16(Op::LOCAL_GET, exc_tmp);
        let msg_k = self.str_const("message");
        self.emit_u16(Op::STRUCT_GET, msg_k);
        fn_call!(self, "wasm:js-string", "concat", 2);
        let stack_val = self.define_local("__stack_val");
        self.emit_u16(Op::LOCAL_SET, stack_val);

        if self.is_js_profile() {
            // Set stack as non-enumerable using Object.defineProperty
            let define_prop_idx = self.import("ecma:object", "defineProperty");
            self.emit_u16(Op::LOCAL_GET, exc_tmp);
            inst!(self, core_wasm::string_const, "stack");
            common::dict::emit_new(&mut self.chunks, self.current, line);
            inst!(self, core_wasm::dup);
            self.emit_u16(Op::LOCAL_GET, stack_val);
            let val_key = self.str_const("value");
            self.emit_u16(Op::STRUCT_SET, val_key);
            self.emit(Op::DROP);
            inst!(self, core_wasm::dup);
            inst!(self, core_wasm::bool_const, false);
            let enum_key = self.str_const("enumerable");
            self.emit_u16(Op::STRUCT_SET, enum_key);
            self.emit(Op::DROP);
            self.emit_host_call(define_prop_idx, 3);
            self.emit(Op::DROP);
        } else {
            self.emit_u16(Op::LOCAL_GET, exc_tmp);
            self.emit_u16(Op::LOCAL_GET, stack_val);
            let stack_key = self.str_const("stack");
            self.emit_u16(Op::STRUCT_SET, stack_key);
            self.emit(Op::DROP);
        }

        if self.is_js_profile() {
            for name in Self::js_error_instanceof_chain(type_name) {
                common::classes::emit_instanceof_chain(
                    &mut self.chunks,
                    self.current,
                    exc_tmp,
                    name,
                    line,
                );
            }
        }

        self.emit_u16(Op::LOCAL_GET, exc_tmp);
        Ok(())
    }

    pub(super) fn emit_js_exception_ctor_value(
        &mut self,
        type_name: &str,
        args: &[&Expression],
    ) -> Result<(), String> {
        if type_name.trim() == "AggregateError" {
            if let Some(msg_arg) = args.get(1) {
                self.compile_expr(msg_arg)?;
            } else {
                self.emit_const(Value::String(Arc::from("")));
            }
            self.emit_js_exception_ctor_from_message_value(type_name)?;

            let exc_tmp = self.define_local("__agg_exc_tmp");
            self.emit_u16(Op::LOCAL_SET, exc_tmp);

            self.emit_u16(Op::LOCAL_GET, exc_tmp);
            if let Some(errors_arg) = args.first() {
                self.compile_expr(errors_arg)?;
            } else {
                common::collections::emit_array_new(&mut self.chunks, self.current, 0, self.line);
            }
            let errors_key = self.str_const("errors");
            self.emit_u16(Op::STRUCT_SET, errors_key);
            self.emit(Op::DROP);

            if let Some(opts_arg) = args.get(2) {
                self.emit_u16(Op::LOCAL_GET, exc_tmp);
                self.compile_expr(opts_arg)?;
                let cause_key = self.str_const("cause");
                self.emit_u16(Op::STRUCT_GET, cause_key);
                let cause_val = self.define_local("__agg_cause_val");
                self.emit_u16(Op::LOCAL_SET, cause_val);
                self.emit_u16(Op::LOCAL_GET, exc_tmp);
                self.emit_u16(Op::LOCAL_GET, cause_val);
                self.emit_u16(Op::STRUCT_SET, cause_key);
                self.emit(Op::DROP);
            }

            self.emit_u16(Op::LOCAL_GET, exc_tmp);
            return Ok(());
        }

        if let Some(msg_arg) = args.first() {
            self.compile_expr(msg_arg)?;
        } else {
            self.emit_const(Value::String(Arc::from("")));
        }
        self.emit_js_exception_ctor_from_message_value(type_name)?;

        if let Some(opts_arg) = args.get(1) {
            let exc_tmp = self.define_local("__exc_with_opts");
            self.emit_u16(Op::LOCAL_SET, exc_tmp);
            // Evaluate the options object once and stash it.
            self.compile_expr(opts_arg)?;
            let opts_tmp = self.define_local("__exc_opts");
            self.emit_u16(Op::LOCAL_SET, opts_tmp);
            // Copy `cause` and `code` from the options object onto the
            // exception so the canonical exception shape carries them
            // uniformly. `new Error(msg, {cause})` (JS/ECMA), PHP's
            // walker-normalized `new Exception(msg, {code, cause})`, etc. all
            // route here, so a PHP-thrown exception's cause/code are visible
            // to a JS/Python catcher and vice-versa. A missing key resolves to
            // Undefined (resolve_property never traps); the language getters
            // apply their own defaults (PHP getPrevious()→null, getCode()→0).
            for key in ["cause", "code"] {
                self.emit_u16(Op::LOCAL_GET, exc_tmp);
                self.emit_u16(Op::LOCAL_GET, opts_tmp);
                let k = self.str_const(key);
                self.emit_u16(Op::STRUCT_GET, k);
                self.emit_u16(Op::STRUCT_SET, k);
                self.emit(Op::DROP);
            }
            self.emit_u16(Op::LOCAL_GET, exc_tmp);
        }
        Ok(())
    }

    pub(crate) fn emit_generator_control_packet_from_stack(&mut self, op: &str) {
        let value_slot = self.define_local("__gen_control_value");
        self.emit_u16(Op::LOCAL_SET, value_slot);

        let line = self.line;
        common::dict::emit_new(&mut self.chunks, self.current, line);

        inst!(self, core_wasm::dup);
        self.emit_const(Value::Bool(true));
        let marker_key = self.str_const("__vybe_generator_control");
        self.emit_u16(Op::STRUCT_SET, marker_key);
        self.emit(Op::DROP);

        inst!(self, core_wasm::dup);
        self.emit_const(Value::String(Arc::from(op)));
        let op_key = self.str_const("op");
        self.emit_u16(Op::STRUCT_SET, op_key);
        self.emit(Op::DROP);

        inst!(self, core_wasm::dup);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        let value_key = self.str_const("value");
        self.emit_u16(Op::STRUCT_SET, value_key);
        self.emit(Op::DROP);
    }

    pub(crate) fn reorder_named_args_with_signatures(
        &self,
        args: &[Argument],
        signatures: &[CallSignature],
    ) -> Vec<Argument> {
        if !args.iter().any(|arg| arg.name.is_some()) {
            return args.to_vec();
        }

        for signature in signatures {
            let mut slots: Vec<Option<Argument>> = vec![None; signature.param_names.len()];
            let mut next_positional = 0usize;
            let mut valid = true;

            for arg in args {
                if arg.spread {
                    valid = false;
                    break;
                }

                let target_index = if let Some(name) = arg.name.as_deref() {
                    signature
                        .param_names
                        .iter()
                        .position(|param_name| param_name.eq_ignore_ascii_case(name))
                } else {
                    while next_positional < slots.len() && slots[next_positional].is_some() {
                        next_positional += 1;
                    }
                    Some(next_positional)
                };
                let Some(index) = target_index else {
                    valid = false;
                    break;
                };
                if index >= slots.len() || slots[index].is_some() {
                    valid = false;
                    break;
                }

                let mut ordered = arg.clone();
                ordered.name = None;
                slots[index] = Some(ordered);

                if arg.name.is_none() {
                    next_positional = index + 1;
                }
            }

            if !valid {
                continue;
            }

            if slots.iter().take(signature.min_arity).any(Option::is_none) {
                continue;
            }

            return slots
                .into_iter()
                .map(|arg| arg.unwrap_or_else(|| Argument::positional(Expression::null())))
                .collect();
        }

        args.to_vec()
    }

    fn reorder_named_call_args(&self, callee: &Expression, args: &[Argument]) -> Vec<Argument> {
        if !args.iter().any(|arg| arg.name.is_some()) {
            return args.to_vec();
        }

        let signatures = match &callee.kind {
            ExprKind::Ident(name) => self.function_signatures.get(&self.canon(name)),
            ExprKind::Member { field, .. } => self.function_signatures.get(&self.canon(field)),
            _ => None,
        };

        signatures
            .map(|signatures| self.reorder_named_args_with_signatures(args, signatures))
            .unwrap_or_else(|| args.to_vec())
    }

    // ════════════════════════════════════════════════════════════════════════
    // Call compilation
    // ════════════════════════════════════════════════════════════════════════

    fn try_compile_go_map_has_call(
        &mut self,
        callee: &Expression,
        args: &[Argument],
    ) -> Result<bool, String> {
        if self.profile.name != "go" || args.len() != 2 {
            return Ok(false);
        }
        let ExprKind::Ident(name) = &callee.kind else {
            return Ok(false);
        };
        if name != "__go_map_has" {
            return Ok(false);
        }

        self.compile_expr(&args[0].value)?;
        let map_slot = self.define_local("__go_map_has_obj");
        self.emit_u16(Op::LOCAL_SET, map_slot);

        self.emit_u16(Op::LOCAL_GET, map_slot);
        self.emit(Op::REF_IS_NULL);
        let line = self.line;
        self.chunk().emit_if_value(line);
        inst!(self, core_wasm::bool_const, false);
        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, map_slot);
        self.compile_expr(&args[1].value)?;
        let line = self.line;
        common::dict::emit_method_has(&mut self.chunks, self.current, line);
        self.chunk().emit_end(line);
        Ok(true)
    }
    pub(super) fn compile_call(
        &mut self,
        callee: &Expression,
        args: &[Argument],
    ) -> Result<(), String> {
        let reordered_args;
        let args = if args.iter().any(|arg| arg.name.is_some()) {
            reordered_args = self.reorder_named_call_args(callee, args);
            reordered_args.as_slice()
        } else {
            args
        };
        let arg_exprs: Vec<&Expression> = args.iter().map(|a| &a.value).collect();

        if self.try_compile_js_iterator_from_generator_take_to_array(callee, args)? {
            return Ok(());
        }

        if self.try_compile_go_map_has_call(callee, args)? {
            return Ok(());
        }

        if self.is_js_profile() {
            if let ExprKind::Ident(name) = &callee.kind {
                if name == "__js_dynamic_import" {
                    let Some(path) = args.first().and_then(|arg| match &arg.value.kind {
                        ExprKind::Lit(Literal::Str(path)) => Some(path.clone()),
                        _ => None,
                    }) else {
                        return Err(
                            "Dynamic import currently requires a string literal module specifier"
                                .into(),
                        );
                    };

                    let normalized_module = self
                        .profile
                        .bare_module_aliases
                        .get(path.as_str())
                        .cloned()
                        .unwrap_or(path);
                    if !super::is_host_specifier(&normalized_module) {
                        self.emit_const(Value::String(Arc::from(format!(
                            "Cannot find module '{}'",
                            normalized_module
                        ))));
                        self.emit_js_exception_ctor_from_message_value("TypeError")?;
                        let reject_idx = self.import("ecma:promise", "reject");
                        self.emit_host_call(reject_idx, 1);
                        return Ok(());
                    }

                    let alias = js_dynamic_import_alias(&normalized_module);
                    self.host_namespace_aliases
                        .insert(self.canon(&alias), normalized_module);
                    self.emit_var_get(&alias);
                    let resolve_idx = self.import("ecma:promise", "resolve");
                    self.emit_host_call(resolve_idx, 1);
                    return Ok(());
                }
                if name == "String" {
                    if let Some(arg) = args.first() {
                        self.compile_expr(&arg.value)?;
                        let idx = self.import("ecma:string", "String");
                        self.emit_host_call(idx, 1);
                    } else {
                        self.emit_const(Value::String(Arc::from("")));
                    }
                    return Ok(());
                }
                if name == "Number" {
                    if let Some(arg) = args.first() {
                        self.compile_expr(&arg.value)?;
                        let arg_slot = self.define_local("__js_number_arg");
                        self.emit_u16(Op::LOCAL_SET, arg_slot);
                        self.emit_u16(Op::LOCAL_GET, arg_slot);
                        fn_call!(self, "ecma:value", "typeof", 1);
                        self.emit_const(Value::String(Arc::from("symbol")));
                        {
                            let line = self.line;
                            crate::emitter::ops::emit_dyn_eq(self.chunk(), line);
                        };
                        let line = self.line;
                        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                        self.chunk().emit_if(line);
                        self.emit_const(Value::String(Arc::from(
                            "Cannot convert a Symbol value to a number",
                        )));
                        self.emit_js_exception_ctor_from_message_value("TypeError")?;
                        let line = self.line;
                        common::errors::emit_throw(self.chunk(), line);
                        self.chunk().emit_end(line);
                        self.emit_u16(Op::LOCAL_GET, arg_slot);
                        let idx = self.import("ecma:number", "Number");
                        self.emit_host_call(idx, 1);
                    } else {
                        self.emit_const(Value::F64(0.0));
                    }
                    return Ok(());
                }
            }
        }

        if let ExprKind::Member {
            object,
            field,
            null_safe,
        } = &callee.kind
        {
            if let Some(text) = self.resolve_reflection_string_member_expr(object) {
                let rewritten = Expression::new(ExprKind::Member {
                    object: Box::new(Expression::string(&text)),
                    field: field.clone(),
                    null_safe: *null_safe,
                });
                return self.compile_call(&rewritten, args);
            }
            if !null_safe
                && field.eq_ignore_ascii_case("Deconstruct")
                && args.iter().all(|arg| arg.by_ref)
            {
                if let ExprKind::Call {
                    callee: inner_callee,
                    args: inner_args,
                    ..
                } = &object.kind
                {
                    if let Some(arity) = self.multi_return_arity_for_callee(inner_callee) {
                        if arity as usize == args.len() {
                            self.compile_call(inner_callee, inner_args)?;
                            for out_arg in args.iter().rev() {
                                if let ExprKind::Ident(name) = &out_arg.value.kind {
                                    if name.starts_with("__discard_") {
                                        self.emit(Op::DROP);
                                        continue;
                                    }
                                }
                                self.compile_assign_target(&out_arg.value)?;
                            }
                            self.emit(Op::NULL);
                            return Ok(());
                        }
                    }
                }
            }
        }

        if self.try_compile_dotnet_case_insensitive_collection_call(callee, args)? {
            return Ok(());
        }
        if self.try_compile_dotnet_delegate_call(callee, args)? {
            return Ok(());
        }
        if self.try_compile_dotnet_numeric_try_parse(callee, args)? {
            return Ok(());
        }
        if self.try_compile_dotnet_dictionary_try_get_value(callee, args)? {
            return Ok(());
        }
        if self.try_compile_dotnet_formatted_tostring(callee, args)? {
            return Ok(());
        }
        if self.try_compile_dotnet_guid_try_parse(callee, args)? {
            return Ok(());
        }
        if self.try_compile_dotnet_enum_call(callee, args)? {
            return Ok(());
        }
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            if field == "call" {
                if let ExprKind::Member {
                    object: to_string_target,
                    field: to_string_field,
                    ..
                } = &object.kind
                {
                    if to_string_field == "toString" {
                        if let ExprKind::Member {
                            object: prototype_target,
                            field: prototype_field,
                            ..
                        } = &to_string_target.kind
                        {
                            if prototype_field == "prototype" {
                                if matches!(&prototype_target.kind, ExprKind::Ident(name) if name == "Object")
                                {
                                    let idx = self.import("ecma:object", "toString");
                                    if let Some(arg) = args.first() {
                                        self.compile_expr(&arg.value)?;
                                        self.emit_host_call(idx, 1);
                                    } else {
                                        self.emit(Op::NULL);
                                        self.emit_host_call(idx, 1);
                                    }
                                    return Ok(());
                                }
                            }
                        }
                    }
                }
            }
        }
        if self.try_compile_dotnet_zero_arg_tostring(callee, args)? {
            return Ok(());
        }
        if self.try_compile_dotnet_attribute_reflection_call(callee, args)? {
            return Ok(());
        }

        if self.is_python_profile() {
            if let ExprKind::Ident(name) = &callee.kind {
                if name == "dict" {
                    let line = self.line;
                    common::dict::emit_new(&mut self.chunks, self.current, line);

                    if args.iter().all(|arg| arg.name.is_some()) {
                        for arg in args {
                            let key = arg.name.as_ref().unwrap();
                            inst!(self, core_wasm::dup);
                            self.compile_expr(&arg.value)?;
                            let key_idx = self.str_const(key);
                            self.emit_u16(Op::STRUCT_SET, key_idx);
                            self.emit(Op::DROP);

                            inst!(self, core_wasm::dup);
                            let keys_key = self.str_const("__keys");
                            self.emit_u16(Op::STRUCT_GET, keys_key);
                            self.emit_const(Value::String(Arc::from(key.as_str())));
                            common::collections::emit_push(&mut self.chunks, self.current, line);
                            self.emit(Op::DROP);
                        }
                        return Ok(());
                    }

                    if args.len() == 1 && args[0].name.is_none() && !args[0].spread {
                        if let ExprKind::Array(elements) = &args[0].value.kind {
                            for element in elements {
                                let ExprKind::Tuple(items) = &element.value.kind else {
                                    continue;
                                };
                                if items.len() != 2 {
                                    continue;
                                }

                                inst!(self, core_wasm::dup);
                                self.compile_expr(&items[0])?;
                                let key_tmp = self.define_local("__py_dict_ctor_key");
                                inst!(self, core_wasm::dup);
                                self.emit_u16(Op::LOCAL_SET, key_tmp);
                                self.compile_expr(&items[1])?;
                                common::collections::emit_set(&mut self.chunks, self.current, line);
                                self.emit(Op::DROP);

                                inst!(self, core_wasm::dup);
                                let keys_key = self.str_const("__keys");
                                self.emit_u16(Op::STRUCT_GET, keys_key);
                                self.emit_u16(Op::LOCAL_GET, key_tmp);
                                common::collections::emit_push(
                                    &mut self.chunks,
                                    self.current,
                                    line,
                                );
                                self.emit(Op::DROP);
                            }
                            return Ok(());
                        }
                    }

                    if args.is_empty() {
                        return Ok(());
                    }
                }
            }
        }

        if self.is_php_profile() {
            if let ExprKind::Ident(name) = &callee.kind {
                if name.eq_ignore_ascii_case("spl_autoload_register") {
                    let receiver_idx = self.str_const("__php_autoload_callback_receiver");
                    if let Some(callback) = args.first() {
                        match &callback.value.kind {
                            ExprKind::Array(elements)
                                if elements.len() == 2
                                    && elements.iter().all(|element| element.key.is_none()) =>
                            {
                                let ExprKind::Lit(Literal::Str(class_name)) =
                                    &elements[0].value.kind
                                else {
                                    inst!(self, core_wasm::undefined);
                                    self.emit_u16(Op::GLOBAL_SET, receiver_idx);
                                    self.compile_php_autoload_callable_ref(&callback.value)?;
                                    let global_idx = self.str_const("__php_autoload_callback");
                                    self.emit_u16(Op::GLOBAL_SET, global_idx);
                                    for arg in args.iter().skip(1) {
                                        self.compile_expr(&arg.value)?;
                                        self.emit(Op::DROP);
                                    }
                                    self.emit_const(Value::Bool(true));
                                    return Ok(());
                                };
                                let ExprKind::Lit(Literal::Str(method_name)) =
                                    &elements[1].value.kind
                                else {
                                    inst!(self, core_wasm::undefined);
                                    self.emit_u16(Op::GLOBAL_SET, receiver_idx);
                                    self.compile_php_autoload_callable_ref(&callback.value)?;
                                    let global_idx = self.str_const("__php_autoload_callback");
                                    self.emit_u16(Op::GLOBAL_SET, global_idx);
                                    for arg in args.iter().skip(1) {
                                        self.compile_expr(&arg.value)?;
                                        self.emit(Op::DROP);
                                    }
                                    self.emit_const(Value::Bool(true));
                                    return Ok(());
                                };

                                if let Some(class_global) =
                                    self.resolve_php_autoload_callback_class_global(class_name)
                                {
                                    let class_idx = self.str_const(&class_global);
                                    self.emit_u16(Op::GLOBAL_GET, class_idx);
                                    inst!(self, core_wasm::dup);
                                    self.emit_u16(Op::GLOBAL_SET, receiver_idx);
                                    let method_idx = self.str_const(&self.canon(method_name));
                                    self.emit_u16(Op::STRUCT_GET, method_idx);
                                } else {
                                    inst!(self, core_wasm::undefined);
                                    self.emit_u16(Op::GLOBAL_SET, receiver_idx);
                                    self.compile_php_autoload_callable_ref(&callback.value)?;
                                }
                            }
                            _ => {
                                inst!(self, core_wasm::undefined);
                                self.emit_u16(Op::GLOBAL_SET, receiver_idx);
                                self.compile_php_autoload_callable_ref(&callback.value)?;
                            }
                        }
                    } else {
                        inst!(self, core_wasm::undefined);
                        self.emit_u16(Op::GLOBAL_SET, receiver_idx);
                        inst!(self, core_wasm::undefined);
                    }
                    let global_idx = self.str_const("__php_autoload_callback");
                    self.emit_u16(Op::GLOBAL_SET, global_idx);

                    for arg in args.iter().skip(1) {
                        self.compile_expr(&arg.value)?;
                        self.emit(Op::DROP);
                    }

                    self.emit_const(Value::Bool(true));
                    return Ok(());
                }

                if name.eq_ignore_ascii_case("spl_autoload_unregister") {
                    for arg in args {
                        self.compile_expr(&arg.value)?;
                        self.emit(Op::DROP);
                    }

                    inst!(self, core_wasm::undefined);
                    let receiver_idx = self.str_const("__php_autoload_callback_receiver");
                    self.emit_u16(Op::GLOBAL_SET, receiver_idx);
                    inst!(self, core_wasm::undefined);
                    let global_idx = self.str_const("__php_autoload_callback");
                    self.emit_u16(Op::GLOBAL_SET, global_idx);
                    self.emit_const(Value::Bool(true));
                    return Ok(());
                }

                if name == "compact" {
                    let line = self.line;
                    common::collections::emit_map_new(&mut self.chunks, self.current, line);
                    for arg in args {
                        let ExprKind::Lit(Literal::Str(var_name)) = &arg.value.kind else {
                            self.emit(Op::NULL);
                            return Ok(());
                        };
                        let php_var_name = format!("${}", var_name);
                        inst!(self, core_wasm::dup);
                        self.emit_const(Value::String(Arc::from(var_name.as_str())));
                        self.emit_var_get(&php_var_name);
                        common::collections::emit_set(&mut self.chunks, self.current, line);
                        self.emit(Op::DROP);
                    }
                    return Ok(());
                }

                if name == "extract" && arg_exprs.len() == 1 {
                    if let ExprKind::Array(elements) = &arg_exprs[0].kind {
                        let mut count = 0i64;
                        for elem in elements {
                            let Some(key_expr) = &elem.key else {
                                continue;
                            };
                            let bind_name = match &key_expr.kind {
                                ExprKind::Lit(Literal::Str(s)) => format!("${}", s),
                                ExprKind::Lit(Literal::Int(n)) => format!("${}", n),
                                _ => continue,
                            };
                            self.compile_expr(&elem.value)?;
                            self.emit_var_set(&bind_name);
                            count += 1;
                        }
                        self.emit_const(Value::I64(count));
                        return Ok(());
                    }

                    let mut binding_names = std::collections::BTreeSet::new();
                    for local in &self.scope().locals {
                        if local.name.starts_with('$') && !local.name.starts_with("$__") {
                            binding_names.insert(local.name.clone());
                        }
                    }
                    for global in &self.defined_globals {
                        if global.starts_with('$') && !global.starts_with("$__") {
                            binding_names.insert(global.clone());
                        }
                    }

                    if !binding_names.is_empty() {
                        let map_slot = self.define_local("__php_extract_map");
                        self.compile_expr(&arg_exprs[0])?;
                        self.emit_u16(Op::LOCAL_SET, map_slot);

                        let count_slot = self.define_local("__php_extract_count");
                        self.emit_const(Value::I64(0));
                        self.emit_u16(Op::LOCAL_SET, count_slot);

                        for bind_name in binding_names {
                            let key_name =
                                bind_name.strip_prefix('$').unwrap_or(bind_name.as_str());
                            self.emit_u16(Op::LOCAL_GET, map_slot);
                            self.emit_const(Value::String(Arc::from(key_name)));
                            let line = self.line;
                            common::collections::emit_get(&mut self.chunks, self.current, line);
                            let value_slot = self.define_local("__php_extract_value");
                            self.emit_u16(Op::LOCAL_SET, value_slot);
                            self.emit_u16(Op::LOCAL_GET, value_slot);
                            self.emit(Op::REF_IS_NULL);
                            let line = self.line;
                            self.chunk().emit_if(line);
                            self.chunk().emit_else(line);
                            self.emit_u16(Op::LOCAL_GET, value_slot);
                            self.emit_var_set(&bind_name);
                            self.emit_u16(Op::LOCAL_GET, count_slot);
                            self.emit_const(Value::I64(1));
                            {
                                let line = self.line;
                                crate::emitter::ops::emit_dyn_add(self.chunk(), line);
                            };
                            self.emit_u16(Op::LOCAL_SET, count_slot);
                            self.chunk().emit_end(line);
                        }

                        self.emit_u16(Op::LOCAL_GET, count_slot);
                        return Ok(());
                    }
                }
            }
        }

        // ── super(args) → call parent constructor, store result as this ──
        if let ExprKind::Super = &callee.kind {
            if let Some(ref class_name) = self.current_class.clone() {
                if let Some(parent_name) = self
                    .pending_classes
                    .get(class_name.as_str())
                    .and_then(|pc| pc.parent.clone())
                {
                    // §13.3.7.2 (JS): super() may only run once — a second
                    // call sees this_slot already initialized and throws a
                    // ReferenceError.
                    if let Some((ctx_chunk, ctx_slot)) = self.js_derived_ctor_ctx {
                        if ctx_chunk == self.current {
                            let l = self.line;
                            common::classes::emit_super_once_guard(self.chunk(), ctx_slot, l);
                        }
                    }
                    if !self.shadows_builtin_type(&parent_name)
                        && common::errors::is_exception_type(&parent_name)
                    {
                        self.emit_js_exception_ctor_value(&parent_name, &arg_exprs)?;
                        let self_kw = self.profile.self_keyword.clone();
                        if let Some(slot) = self
                            .scope()
                            .resolve(&self_kw)
                            .or_else(|| self.scope().resolve_ci(&self_kw))
                        {
                            inst!(self, core_wasm::dup);
                            self.emit_u16(Op::LOCAL_SET, slot);
                        }
                        return Ok(());
                    }
                    self.emit_var_get(&parent_name);
                    for a in &arg_exprs {
                        self.compile_expr(a)?;
                    }
                    self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                    // Store result as this
                    let self_kw = self.profile.self_keyword.clone();
                    if let Some(slot) = self
                        .scope()
                        .resolve(&self_kw)
                        .or_else(|| self.scope().resolve_ci(&self_kw))
                    {
                        inst!(self, core_wasm::dup);
                        self.emit_u16(Op::LOCAL_SET, slot);
                    }
                    return Ok(());
                }
            }
            // No parent — emit null
            self.emit(Op::NULL);
            return Ok(());
        }

        // ── super.method(args) — static class dispatch ───────────────
        //
        // Resolve the parent class statically at compile time. Inside
        // `class C extends B`, `super.method()` always means B's
        // method (regardless of the runtime instance type) — the spec
        // says super uses [[HomeObject]].[[Prototype]], NOT the
        // instance's prototype chain. Multi-level inheritance (C → B
        // → A) needs B.method when called from C and A.method when
        // called from B; the previous `this.__base_method` lookup
        // collided across levels (C overwriting B's slot) and caused
        // an infinite loop on C's super chain.
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            if matches!(&object.kind, ExprKind::Super) {
                let canon_field = self.canon(field);
                let class_name = self.current_class.clone();
                let parent_name = class_name
                    .as_ref()
                    .and_then(|cn| self.pending_classes.get(cn.as_str()))
                    .and_then(|pc| pc.parent.clone());
                let self_kw = self.profile.self_keyword.clone();
                let self_slot = self
                    .scope()
                    .resolve(&self_kw)
                    .or_else(|| self.scope().resolve_ci(&self_kw));

                if let Some(parent) = parent_name {
                    // Look up parent class via emit_var_get so closure-
                    // captured parents (mixin pattern: `(Base) => class
                    // extends Base`) resolve through the upvalue scope.
                    self.emit_var_get(&parent);
                    let method_idx = self.str_const(&canon_field);
                    self.emit_u16(Op::STRUCT_GET, method_idx);

                    if self.is_js_profile() {
                        let saved_js_this = self.save_js_this("__js_prev_this_super_method");
                        if let Some(slot) = self_slot {
                            self.emit_u16(Op::LOCAL_GET, slot);
                        } else {
                            let js_this = self.str_const("__js_this");
                            self.emit_u16(Op::GLOBAL_GET, js_this);
                        }
                        self.set_js_this_from_stack();
                        for a in &arg_exprs {
                            self.compile_expr(a)?;
                        }
                        self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                        let result_slot = self.define_local("__js_super_method_result");
                        self.emit_u16(Op::LOCAL_SET, result_slot);
                        self.restore_js_this(saved_js_this);
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                    } else {
                        // Typed-language method ABI passes receiver as arg0.
                        if let Some(slot) = self_slot {
                            self.emit_u16(Op::LOCAL_GET, slot);
                        } else {
                            self.emit(Op::NULL);
                        }
                        for a in &arg_exprs {
                            self.compile_expr(a)?;
                        }
                        self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                    }
                    return Ok(());
                }

                // Pascal / VB / C# allow `inherited Foo` / `MyBase.Foo` in a
                // root class even when there is no parent implementation. Treat
                // it as a no-op instead of falling through to the generic member
                // call pipeline and recursing back into the current method.
                self.emit(Op::NULL);
                return Ok(());
            }
        }

        // ── Debug intrinsic: __debug_dump(obj) ──────────────────────
        // Available in all languages. Serialises each arg via JSON.stringify
        // then emits it to wasi:logging/logging.log.
        if let ExprKind::Ident(name) = &callee.kind {
            if name == "__debug_dump" {
                let stringify_idx = self.import("ecma:json", "stringify");
                let log_idx = self.import("wasi:logging/logging", "log");
                for a in &arg_exprs {
                    self.compile_expr(a)?;
                    self.emit_host_call(stringify_idx, 1);
                    self.emit_host_call(log_idx, 1);
                }
                return Ok(());
            }

            let canon = self.canon(name);
            let shadows_builtin_exception = self.defined_functions.contains(&canon)
                || self.defined_classes.contains(&canon)
                || self.defined_globals.contains(&canon)
                || (!self.case_sensitive
                    && (self
                        .defined_functions
                        .iter()
                        .any(|g| g.eq_ignore_ascii_case(name))
                        || self
                            .defined_classes
                            .iter()
                            .any(|g| g.eq_ignore_ascii_case(name))
                        || self
                            .defined_globals
                            .iter()
                            .any(|g| g.eq_ignore_ascii_case(name))));
            if !shadows_builtin_exception && common::errors::is_exception_type(name) {
                self.emit_js_exception_ctor_value(name, &arg_exprs)?;
                return Ok(());
            }
        }

        // ── Typed static-field receiver: counts.ContainsKey(...) ─────
        // Static fields can carry type hints too. Resolve them here so
        // class-level typed state uses the same shared .NET surface as
        // locals with type annotations.
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            let class_name = resolve_receiver_type_hint(self, object);
            if let Some(class_name) = class_name {
                if self
                    .resolve_pending_class_name_for_type_hint(&class_name)
                    .is_some()
                {
                    // User-defined classes win over shared .NET surface names
                    // like `Stack`, `Queue`, or `Dictionary`.
                } else {
                    let class_name = Self::normalize_type_hint(&class_name);
                    let surface = common::dotnet::surface();
                    if let Some(target) =
                        surface.lookup_instance_method(&class_name, field, arg_exprs.len() as u8)
                    {
                        if matches!(&target, common::dotnet::InstanceMethodTarget::Common { emit, .. } if emit == "collections.sort")
                            && arg_exprs.is_empty()
                            && !self.is_js_profile()
                        {
                            let sort_global = self.str_const("__vybe_sort_with_comparator");
                            self.emit_u16(Op::GLOBAL_GET, sort_global);
                            self.compile_expr(object)?;
                            self.compile_lambda(
                                &[
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
                                &LambdaBody::Expr(Box::new(Expression::new(ExprKind::Ternary {
                                    cond: Box::new(Expression::new(ExprKind::Binary {
                                        op: BinOp::Lt,
                                        left: Box::new(Expression::ident("left")),
                                        right: Box::new(Expression::ident("right")),
                                    })),
                                    then: Box::new(Expression::new(ExprKind::Lit(Literal::Int(
                                        -1,
                                    )))),
                                    else_: Box::new(Expression::new(ExprKind::Ternary {
                                        cond: Box::new(Expression::new(ExprKind::Binary {
                                            op: BinOp::Gt,
                                            left: Box::new(Expression::ident("left")),
                                            right: Box::new(Expression::ident("right")),
                                        })),
                                        then: Box::new(Expression::new(ExprKind::Lit(
                                            Literal::Int(1),
                                        ))),
                                        else_: Box::new(Expression::new(ExprKind::Lit(
                                            Literal::Int(0),
                                        ))),
                                    })),
                                }))),
                                &[],
                            )?;
                            self.emit_u8(Op::CALL_REF, 2);
                            return Ok(());
                        }

                        if matches!(&target, common::dotnet::InstanceMethodTarget::Common { emit, .. } if emit == "dotnet.array_sort")
                            && arg_exprs.len() == 1
                            && !self.is_js_profile()
                            && class_name.rsplit('.').next().is_some_and(|name| {
                                name.eq_ignore_ascii_case("List")
                                    || name.eq_ignore_ascii_case("ArrayList")
                            })
                            && matches!(&arg_exprs[0].kind, ExprKind::Lambda { params, .. } if params.len() == 2)
                        {
                            let sort_global = self.str_const("__vybe_sort_with_comparator");
                            self.emit_u16(Op::GLOBAL_GET, sort_global);
                            self.compile_expr(object)?;
                            self.compile_expr(&arg_exprs[0])?;
                            self.emit_u8(Op::CALL_REF, 2);
                            return Ok(());
                        }

                        self.compile_expr(object)?;
                        for a in &arg_exprs {
                            self.compile_expr(a)?;
                        }
                        let total_argc = (arg_exprs.len() + 1) as u8;
                        match target {
                            common::dotnet::InstanceMethodTarget::Host { module, func, .. } => {
                                let idx = self.import(&module, &func);
                                self.emit_host_call(idx, total_argc);
                            }
                            common::dotnet::InstanceMethodTarget::Common { emit, .. } => {
                                let line = self.line;
                                self.emit_common(&emit, total_argc, line);
                            }
                        }
                        return Ok(());
                    }
                }
            }
        }
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            if resolves_to_static_container_method(self, object, field) {
                self.compile_expr(object)?;
                let obj_tmp = self.define_local("__static_container_obj");
                self.emit_u16(Op::LOCAL_SET, obj_tmp);
                let fn_tmp = self.define_local("__static_container_fn");
                let class_canon = self.canon(&self.flatten_member_chain(object).join("."));
                if self.is_js_profile() && field.starts_with('#') {
                    if let Some(overload) = self.resolve_static_method_overload_for_type(
                        &class_canon,
                        field,
                        &arg_exprs,
                    ) {
                        let line = self.line;
                        self.emit_u16(Op::REF_FUNC, overload.chunk_idx as u16);
                        self.chunk().emit(0, line);
                    } else if let Some(chunk_idx) =
                        self.resolve_unique_static_method_chunk_for_class(&class_canon, field)
                    {
                        let line = self.line;
                        self.emit_u16(Op::REF_FUNC, chunk_idx as u16);
                        self.chunk().emit(0, line);
                    } else {
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        let method_name =
                            self.js_member_storage_name_for_class(&class_canon, field);
                        let method_idx = self.str_const(&method_name);
                        self.emit_u16(Op::STRUCT_GET, method_idx);
                    }
                } else {
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    let method_idx = self.str_const(&self.canon(field));
                    self.emit_u16(Op::STRUCT_GET, method_idx);
                }
                self.emit_u16(Op::LOCAL_SET, fn_tmp);
                if self.profile.name == "csharp" && args.len() == 1 && !args[0].spread {
                    if self
                        .resolve_static_method_overload_for_type(&class_canon, field, &arg_exprs)
                        .is_some_and(|overload| overload.signature.has_rest)
                    {
                        self.emit_variadic_array_call_from_local(fn_tmp, &args[0].value)?;
                        return Ok(());
                    }
                }
                let rest_signature = self
                    .resolve_static_method_overload_for_type(&class_canon, field, &arg_exprs)
                    .map(|overload| overload.signature.clone())
                    .filter(|signature| signature.has_rest)
                    .or_else(|| {
                        self.function_signatures
                            .get(&self.canon(field))
                            .and_then(|signatures| self.select_call_signature(signatures, args))
                            .filter(|signature| signature.has_rest)
                            .cloned()
                    });
                if let Some(signature) = rest_signature.as_ref() {
                    self.emit_known_rest_call_from_local(
                        fn_tmp,
                        if self.profile.name == "php" {
                            Some(obj_tmp)
                        } else {
                            None
                        },
                        args,
                        signature,
                    )?;
                } else if self.profile.name == "php" {
                    let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                    for (index, arg) in arg_exprs.iter().enumerate() {
                        self.compile_expr(arg)?;
                        let arg_slot =
                            self.define_local(&format!("__static_container_php_arg_{}", index));
                        self.emit_u16(Op::LOCAL_SET, arg_slot);
                        arg_slots.push(arg_slot);
                    }
                    self.emit_call_ref_with_arg_slots(fn_tmp, Some(obj_tmp), &arg_slots);
                } else {
                    let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                    for (index, arg) in arg_exprs.iter().enumerate() {
                        self.compile_expr(arg)?;
                        let arg_slot =
                            self.define_local(&format!("__static_container_arg_{}", index));
                        self.emit_u16(Op::LOCAL_SET, arg_slot);
                        arg_slots.push(arg_slot);
                    }
                    if self.class_prototype_dispatch() {
                        // A static method call binds `this` to the class
                        // object it was fetched from (`this.name` inside
                        // `static describe()` sees the receiving class —
                        // including subclasses inheriting the static).
                        let result_slot = self.define_local("__static_container_result");
                        self.emit_dispatch_and_store_from_arg_slots(
                            fn_tmp,
                            None,
                            Some(obj_tmp),
                            &arg_slots,
                            result_slot,
                        );
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                    } else {
                        self.emit_call_ref_with_arg_slots(fn_tmp, None, &arg_slots);
                    }
                }
                if args.iter().any(|arg| arg.by_ref) {
                    let pack_slot = self.define_local("__static_container_by_ref_pack");
                    self.emit_u16(Op::LOCAL_SET, pack_slot);
                    let mut ref_out_index = 1usize;
                    for arg in args {
                        if !arg.by_ref {
                            continue;
                        }
                        self.emit_u16(Op::LOCAL_GET, pack_slot);
                        self.emit_const(Value::F64(ref_out_index as f64));
                        common::collections::emit_get(&mut self.chunks, self.current, self.line);
                        self.compile_assign_target(&arg.value)?;
                        ref_out_index += 1;
                    }
                    self.emit_u16(Op::LOCAL_GET, pack_slot);
                    self.emit_const(Value::F64(0.0));
                    common::collections::emit_get(&mut self.chunks, self.current, self.line);
                }
                return Ok(());
            }
        }

        // ── Builtin check: Ident("print") ───────────────────────────
        // Skip for user-defined functions: a VB `Function Echo(...)` must
        // dispatch to the user's chunk, not to the cross-language `echo →
        // wasi:cli.log` import shortcut.
        if let ExprKind::Ident(name) = &callee.kind {
            let shadows_builtin = self.defined_functions.contains(name)
                || (!self.case_sensitive
                    && self
                        .defined_functions
                        .iter()
                        .any(|g| g.eq_ignore_ascii_case(name)));
            if !shadows_builtin {
                if self.try_compile_spread_host_builtin(callee, name, args)? {
                    return Ok(());
                }
                if self.try_compile_builtin(name, &arg_exprs)? {
                    return Ok(());
                }
            }
        }

        // ── Builtin check: Member("Console.WriteLine") ─────────────
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            if let ExprKind::Ident(obj_name) = &object.kind {
                // Note: Object.create is handled via the host fn
                // (`ecma:object.create`) so it gets the full ECMA-262
                // §20.1.2.2 behaviour: descriptor second-arg, null
                // prototype gets `toString` etc. stamped as Undefined,
                // and parent properties are copied down for member
                // access. The earlier compiler shortcut here only set
                // `__proto__` and missed both — falling through to
                // `try_compile_builtin` below routes to the host fn.

                let compound = format!("{}.{}", obj_name, field);
                if self.try_compile_spread_host_builtin(callee, &compound, args)? {
                    return Ok(());
                }
                if self.try_compile_builtin(&compound, &arg_exprs)? {
                    return Ok(());
                }

                // ── ESM wildcard namespace member call ──────────────
                //
                // Per ECMA-262 §16.2, a Module Namespace Object is a
                // compile-time binding — `ns.field` resolves statically
                // to the `(module, field)` export. Covers both profile
                // defaults (JS `console` → `wasi:cli`) and user wildcard
                // imports (`import * as cli from "wasi:cli"`). The
                // Linker populated both into `host_namespace_aliases`.
                //
                // Runs AFTER `try_compile_builtin(compound)` so profile
                // builtins with custom emit logic (`Array.from`,
                // `Math.max`) still win on the names they claim.
                let key = self.canon(obj_name);
                if let Some(module) = self.host_namespace_aliases.get(&key).cloned() {
                    let _ = module;
                    let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                    for (index, arg) in arg_exprs.iter().enumerate() {
                        self.compile_expr(arg)?;
                        let arg_slot =
                            self.define_local(&format!("__host_namespace_arg_{}", index));
                        self.emit_u16(Op::LOCAL_SET, arg_slot);
                        arg_slots.push(arg_slot);
                    }
                    self.emit_var_get(obj_name);
                    let namespace_slot = self.define_local("__host_namespace_call_ns");
                    self.emit_u16(Op::LOCAL_SET, namespace_slot);
                    self.emit_u16(Op::LOCAL_GET, namespace_slot);
                    let field_idx = self.str_const(field);
                    self.emit_u16(Op::STRUCT_GET, field_idx);
                    let callee_slot = self.define_local("__host_namespace_call_callee");
                    self.emit_u16(Op::LOCAL_SET, callee_slot);
                    self.emit_call_ref_with_arg_slots(callee_slot, None, &arg_slots);
                    return Ok(());
                }
            }
        }

        if let ExprKind::Member { .. } = &callee.kind {
            let parts = self.flatten_member_chain(callee);
            if parts.len() >= 2 {
                let compound = parts.join(".");
                if self.try_compile_builtin(&compound, &arg_exprs)? {
                    return Ok(());
                }
                // Component model fallback: try dotnet resolver for System.* chains
                if self.try_compile_dotnet_component_call(&parts, &arg_exprs)? {
                    return Ok(());
                }
            }
        }

        // ── Two-level host prefix: `vybe.gui.setProperty(...)` ──────
        //
        // VB / languages without ESM imports reach host functions via
        // a literal namespace chain `<prefix>.<module>.<fn>(args)` where
        // the leading ident is a known host-namespace prefix (`vybe`,
        // `wasi`, `wasm`). Emit as `call_import("<prefix>:<module>",
        // "<fn>", args)` — identical to what JS gets via `import * as
        // gui from "vybe:gui"; gui.setProperty(...)`.
        //
        // Without this, the call falls through to the method-call
        // pattern and injects `vybe.<module>` as a phantom receiver,
        // shifting every argument right by one and silently breaking
        // host functions that don't expect a receiver slot.
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            if let ExprKind::Member {
                object: inner_obj,
                field: inner_field,
                ..
            } = &object.kind
            {
                if let ExprKind::Ident(prefix) = &inner_obj.kind {
                    let prefix_lc = self.canon(prefix);
                    if matches!(prefix_lc.as_str(), "vybe" | "wasi" | "wasm") {
                        let module = format!("{}:{}", prefix_lc, self.canon(inner_field));
                        let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                        for (index, arg) in arg_exprs.iter().enumerate() {
                            self.compile_expr(arg)?;
                            let arg_slot =
                                self.define_local(&format!("__host_prefix_arg_{}", index));
                            self.emit_u16(Op::LOCAL_SET, arg_slot);
                            arg_slots.push(arg_slot);
                        }
                        for slot in &arg_slots {
                            self.emit_u16(Op::LOCAL_GET, *slot);
                        }
                        let idx = self.import(&module, field);
                        self.emit_host_call(idx, arg_exprs.len() as u8);
                        return Ok(());
                    }
                }
            }
        }

        // ── Dotted name resolution FIRST (uses compiler_common::dotnet when use_dotnet) ──
        // Must run before value methods because value methods like "add" would
        // intercept "Controls.Add" which needs special GUI handling.
        if let ExprKind::Member { .. } = &callee.kind {
            let parts = self.flatten_member_chain(callee);
            if parts.len() >= 2 {
                let lower_parts: Vec<String> = parts.iter().map(|s| self.canon(s)).collect();
                let class_parts = &parts[..parts.len() - 1];
                let method_name = parts.last().cloned().unwrap_or_default();

                let mut early_static_class_canon = None;
                if !class_parts.is_empty() {
                    let class_path = class_parts.join(".");
                    let head_name = class_parts.first().map(String::as_str).unwrap_or("");
                    let full_canon = self.canon(&class_path);
                    if self.defined_classes.contains(&full_canon)
                        && self.scope().resolve(head_name).is_none()
                        && self.scope().resolve_ci(head_name).is_none()
                        && self.lookup_var_type_hint(head_name).is_none()
                    {
                        early_static_class_canon = Some(full_canon);
                    }

                    if early_static_class_canon.is_none() && class_parts.len() > 1 {
                        let short_name = class_parts.last().map(String::as_str).unwrap_or("");
                        let short_canon = self.canon(short_name);
                        if self.defined_classes.contains(&short_canon)
                            && self.scope().resolve(short_name).is_none()
                            && self.scope().resolve_ci(short_name).is_none()
                            && self.lookup_var_type_hint(short_name).is_none()
                        {
                            early_static_class_canon = Some(short_canon);
                        }
                    }
                }

                // §20.2.3 (JS): bind/call/apply on a class object resolve
                // through %Function.prototype% unless shadowed by an own
                // static — skip static dispatch so the generic
                // Function.prototype route handles them.
                if self.is_js_profile()
                    && matches!(method_name.as_str(), "bind" | "call" | "apply")
                {
                    if let Some(class_canon) = early_static_class_canon.as_ref() {
                        let canon_field = self.canon(&method_name);
                        let shadowed = self.pending_classes.get(class_canon.as_str()).is_some_and(
                            |pc| pc.static_method_overloads.contains_key(&canon_field),
                        );
                        if !shadowed {
                            early_static_class_canon = None;
                        }
                    }
                }

                if let Some(class_canon) = early_static_class_canon {
                    let cls_idx = self.str_const(&class_canon);
                    self.emit_u16(Op::GLOBAL_GET, cls_idx);
                    let method_canon = self.canon(&method_name);
                    let qualified_method = self.canon(&format!("{}.{}", class_canon, method_name));
                    let method_idx = self.str_const(&method_canon);
                    self.emit_u16(Op::STRUCT_GET, method_idx);
                    let fn_tmp = self
                        .scope()
                        .resolve("__early_static_fn")
                        .unwrap_or_else(|| self.define_local("__early_static_fn"));
                    self.emit_u16(Op::LOCAL_SET, fn_tmp);

                    if let Some(param_modes) = self
                        .function_param_modes
                        .get(&qualified_method)
                        .cloned()
                        .or_else(|| self.function_param_modes.get(&method_canon).cloned())
                    {
                        if param_modes
                            .iter()
                            .any(|mode| matches!(mode, PassBy::Ref | PassBy::Out))
                        {
                            let mut arg_slots = Vec::with_capacity(args.len());
                            for (index, arg) in args.iter().enumerate() {
                                self.compile_ref_aware_call_arg(
                                    arg,
                                    param_modes.get(index).copied().unwrap_or(PassBy::Value),
                                )?;
                                let arg_slot = self
                                    .define_local(&format!("__early_static_call_arg_{}", index));
                                self.emit_u16(Op::LOCAL_SET, arg_slot);
                                arg_slots.push(arg_slot);
                            }

                            self.emit_u16(Op::LOCAL_GET, fn_tmp);
                            for slot in &arg_slots {
                                self.emit_u16(Op::LOCAL_GET, *slot);
                            }
                            self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);

                            let pack_slot = self.define_local("__early_static_ref_call_pack");
                            self.emit_u16(Op::LOCAL_SET, pack_slot);
                            let mut ref_out_index = 1usize;
                            for (index, arg) in args.iter().enumerate() {
                                if !matches!(
                                    param_modes.get(index),
                                    Some(PassBy::Ref | PassBy::Out)
                                ) {
                                    continue;
                                }
                                self.emit_u16(Op::LOCAL_GET, pack_slot);
                                self.emit_const(Value::F64(ref_out_index as f64));
                                common::collections::emit_get(
                                    &mut self.chunks,
                                    self.current,
                                    self.line,
                                );
                                self.compile_assign_target(&arg.value)?;
                                ref_out_index += 1;
                            }
                            self.emit_u16(Op::LOCAL_GET, pack_slot);
                            self.emit_const(Value::F64(0.0));
                            common::collections::emit_get(
                                &mut self.chunks,
                                self.current,
                                self.line,
                            );
                            return Ok(());
                        }
                    }

                    if self.profile.name == "csharp" && args.len() == 1 && !args[0].spread {
                        if self
                            .resolve_static_method_overload_for_type(
                                &class_canon,
                                &method_name,
                                &arg_exprs,
                            )
                            .is_some_and(|overload| overload.signature.has_rest)
                        {
                            self.emit_variadic_array_call_from_local(fn_tmp, &args[0].value)?;
                            return Ok(());
                        }
                    }
                    let rest_signature = self
                        .function_signatures
                        .get(&qualified_method)
                        .and_then(|signatures| self.select_call_signature(signatures, args))
                        .filter(|signature| signature.has_rest)
                        .cloned()
                        .or_else(|| {
                            self.function_signatures
                                .get(&method_canon)
                                .and_then(|signatures| self.select_call_signature(signatures, args))
                                .filter(|signature| signature.has_rest)
                                .cloned()
                        });
                    if let Some(signature) = rest_signature.as_ref() {
                        self.emit_known_rest_call_from_local(fn_tmp, None, args, signature)?;
                    } else {
                        let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                        for (index, arg) in arg_exprs.iter().enumerate() {
                            self.compile_expr(arg)?;
                            let arg_slot =
                                self.define_local(&format!("__early_static_arg_{}", index));
                            self.emit_u16(Op::LOCAL_SET, arg_slot);
                            arg_slots.push(arg_slot);
                        }
                        if self.profile.name == "php" {
                            let cls_idx = self.str_const(&class_canon);
                            self.emit_u16(Op::GLOBAL_GET, cls_idx);
                            let receiver_slot = self.define_local("__early_static_receiver");
                            self.emit_u16(Op::LOCAL_SET, receiver_slot);
                            self.emit_call_ref_with_arg_slots(
                                fn_tmp,
                                Some(receiver_slot),
                                &arg_slots,
                            );
                        } else {
                            self.emit_call_ref_with_arg_slots(fn_tmp, None, &arg_slots);
                        }
                    }
                    if args.iter().any(|arg| arg.by_ref) {
                        let pack_slot = self.define_local("__early_static_by_ref_pack");
                        self.emit_u16(Op::LOCAL_SET, pack_slot);
                        let mut ref_out_index = 1usize;
                        for arg in args {
                            if !arg.by_ref {
                                continue;
                            }
                            self.emit_u16(Op::LOCAL_GET, pack_slot);
                            self.emit_const(Value::F64(ref_out_index as f64));
                            common::collections::emit_get(
                                &mut self.chunks,
                                self.current,
                                self.line,
                            );
                            self.compile_assign_target(&arg.value)?;
                            ref_out_index += 1;
                        }
                        self.emit_u16(Op::LOCAL_GET, pack_slot);
                        self.emit_const(Value::F64(0.0));
                        common::collections::emit_get(&mut self.chunks, self.current, self.line);
                    }
                    return Ok(());
                }

                // Use dotnet resolver when enabled
                if self.profile.namespaces.use_dotnet_resolver {
                    let skip_simple_instance_chain = if lower_parts.len() == 2 {
                        let head = &parts[0];
                        self.has_accessible_local_binding(head)
                            || self.defined_globals.contains(head)
                            || self
                                .defined_globals
                                .iter()
                                .any(|g| g.eq_ignore_ascii_case(head))
                            || self.is_class_field(head)
                            || self.is_class_static_field(head).is_some()
                    } else {
                        false
                    };
                    if skip_simple_instance_chain {
                        // Keep 2-part local/global member calls (`x.Method(...)`) on the
                        // normal instance pipeline; the dotted resolver is for namespace/
                        // static chains and can otherwise short-circuit LINQ-style calls.
                    } else {
                        let dotnet_surface = common::dotnet::surface();
                        let imports = {
                            let mut imp = dotnet_surface.default_imports().to_vec();
                            imp.extend(self.profile.namespaces.extra_imports.clone());
                            imp
                        };
                        let defined_globals = self.defined_globals.clone();
                        let field_set: std::collections::HashSet<String> =
                            if let Some(ref cn) = self.current_class {
                                self.pending_classes
                                    .get(cn.as_str())
                                    .map(|pc| pc.fields.iter().cloned().collect())
                                    .unwrap_or_default()
                            } else {
                                std::collections::HashSet::new()
                            };
                        // `is_local` must recognise top-level variables that
                        // live in `defined_globals` (VB `Dim` at the module
                        // level, JS top-level `var`/`let`), but MUST NOT
                        // match user classes there — those go through
                        // `is_user_type` which returns Unresolved so static
                        // dispatch runs the class ctor path, not a bogus
                        // struct_get chain off the ctor function. The union
                        // (`is_local`) minus (`is_user_type`) gives the
                        // right set of "things you can local_get and
                        // struct_get from".
                        let defined_classes = self.defined_classes.clone();
                        let is_user_class_fn = move |name: &str| -> bool {
                            defined_classes.contains(name)
                                || defined_classes.iter().any(|c| c.eq_ignore_ascii_case(name))
                        };
                        let is_user_class_for_local = is_user_class_fn.clone();
                        let accessible_locals = self
                            .scopes
                            .iter()
                            .flat_map(|scope| scope.locals.iter().map(|local| local.name.clone()))
                            .collect::<Vec<_>>();
                        let ctx = common::dotnet::ResolutionContext {
                            is_local: &|name: &str| {
                                if is_user_class_for_local(name) {
                                    return false;
                                }
                                accessible_locals
                                    .iter()
                                    .any(|local| local == name || local.eq_ignore_ascii_case(name))
                                    || defined_globals.contains(name)
                                    || defined_globals.iter().any(|g| g.eq_ignore_ascii_case(name))
                            },
                            is_class_field: &|name: &str| field_set.contains(name),
                            is_user_type: &is_user_class_fn,
                            imports: &imports,
                        };
                        let refs: Vec<&str> = lower_parts.iter().map(|s| s.as_str()).collect();
                        let resolution = common::dotnet::resolve_dotted_name(&refs, &ctx);

                        match resolution {
                            common::dotnet::DottedResolution::GlobalAccess { name } => {
                                let global_idx = self.str_const(&name);
                                self.emit_u16(Op::GLOBAL_GET, global_idx);
                                for a in &arg_exprs {
                                    self.compile_expr(a)?;
                                }
                                self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                                return Ok(());
                            }
                            common::dotnet::DottedResolution::CommonCall { emit } => {
                                if emit.eq_ignore_ascii_case("dotnet.array_resize")
                                    && args.len() == 2
                                    && args[0].by_ref
                                {
                                    self.compile_expr(&args[0].value)?;
                                    self.compile_expr(&args[1].value)?;
                                    let line = self.line;
                                    self.emit_common(&emit, 2, line);
                                    let resized_slot = self.define_local("__array_resize_value");
                                    self.emit_u16(Op::LOCAL_SET, resized_slot);
                                    self.emit_u16(Op::LOCAL_GET, resized_slot);
                                    self.compile_assign_target(&args[0].value)?;
                                    self.emit(Op::NULL);
                                    return Ok(());
                                }

                                if emit.eq_ignore_ascii_case("dotnet.console_writeline")
                                    && arg_exprs.len() == 1
                                {
                                    self.emit_dotnet_console_arg(arg_exprs[0])?;
                                } else {
                                    for a in &arg_exprs {
                                        self.compile_expr(a)?;
                                    }
                                }
                                let line = self.line;
                                self.emit_common(&emit, arg_exprs.len() as u8, line);
                                return Ok(());
                            }
                            common::dotnet::DottedResolution::HostCall { module, func } => {
                                if self.profile.name == "csharp"
                                    && module.eq_ignore_ascii_case("ecma:number")
                                    && func.eq_ignore_ascii_case("parseInt")
                                    && arg_exprs.len() == 1
                                {
                                    let is_char_like = match &arg_exprs[0].kind {
                                        ExprKind::Lit(Literal::Char(_)) => true,
                                        ExprKind::Ident(name) => {
                                            self.lookup_var_type_hint(name).is_some_and(|hint| {
                                                Self::normalize_type_hint(hint) == "char"
                                            })
                                        }
                                        _ => false,
                                    };
                                    if is_char_like {
                                        self.compile_expr(arg_exprs[0])?;
                                        inst!(self, core_wasm::i32_const, 0);
                                        fn_call!(self, "wasm:js-string", "charCodeAt", 2);
                                        return Ok(());
                                    }
                                }
                                for a in &arg_exprs {
                                    self.compile_expr(a)?;
                                }
                                let idx = self.import(&module, &func);
                                self.emit_host_call(idx, arg_exprs.len() as u8);
                                return Ok(());
                            }
                            common::dotnet::DottedResolution::NamespaceAccess {
                                parts: ns_parts,
                            } => {
                                // If any contiguous sub-window of the chain is a profile namespace
                                // constant (e.g. ["system","math","pi","tostring"] where "math.pi"
                                // is a constant), emit the constant and dispatch remaining as a
                                // value method. Namespace prefix before the constant is discarded.
                                if ns_parts.len() >= 2 {
                                    let mut found_window: Option<(usize, usize)> = None;
                                    'outer: for start in 0..ns_parts.len().saturating_sub(1) {
                                        for end in
                                            ((start + 2)..=ns_parts.len().saturating_sub(0)).rev()
                                        {
                                            if end > ns_parts.len() {
                                                continue;
                                            }
                                            let key = ns_parts[start..end].join(".");
                                            if self.profile.lookup_constant(&key).is_some() {
                                                found_window = Some((start, end));
                                                break 'outer;
                                            }
                                        }
                                    }
                                    if let Some((_const_start, const_end)) = found_window {
                                        let key = ns_parts[_const_start..const_end].join(".");
                                        let cv =
                                            self.profile.lookup_constant(&key).cloned().unwrap();
                                        match &cv {
                                            ConstantValue::Float(f) => {
                                                self.emit_const(Value::F64(*f))
                                            }
                                            ConstantValue::Str(s) => self
                                                .emit_const(Value::String(Arc::from(s.as_str()))),
                                        }
                                        let remaining = ns_parts[const_end..].to_vec();
                                        if let Some(method_name) = remaining.first() {
                                            let argc = arg_exprs.len() as u8;
                                            let def = self
                                                .profile
                                                .lookup_value_method(method_name, argc)
                                                .cloned();
                                            if let Some(def) = def {
                                                for a in &arg_exprs {
                                                    self.compile_expr(a)?;
                                                }
                                                let line = self.line;
                                                match &def.emit {
                                                    BuiltinEmit::HostCall(module, func) => {
                                                        let idx = self.import(module, func);
                                                        self.emit_host_call(
                                                            idx,
                                                            (arg_exprs.len() + 1) as u8,
                                                        );
                                                    }
                                                    BuiltinEmit::Common(name) => {
                                                        let name = name.clone();
                                                        self.emit_common(
                                                            &name,
                                                            (arg_exprs.len() + 1) as u8,
                                                            line,
                                                        );
                                                    }
                                                    BuiltinEmit::Opcode(op_name) => {
                                                        self.emit_named_opcode(op_name);
                                                    }
                                                    _ => {
                                                        // Fallback: STRUCT_GET the method and call_ref
                                                        let idx = self.str_const(method_name);
                                                        self.emit_u16(Op::STRUCT_GET, idx);
                                                        self.emit_u8(
                                                            Op::CALL_REF,
                                                            arg_exprs.len() as u8,
                                                        );
                                                    }
                                                }
                                            } else {
                                                // No value method — STRUCT_GET and call_ref
                                                let idx = self.str_const(method_name);
                                                self.emit_u16(Op::STRUCT_GET, idx);
                                                for a in &arg_exprs {
                                                    self.compile_expr(a)?;
                                                }
                                                self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                                            }
                                        }
                                        return Ok(());
                                    }
                                }

                                if !arg_exprs.is_empty() && ns_parts.len() >= 2 {
                                    let method_name = ns_parts.last().cloned().unwrap_or_default();
                                    let root_idx = self.str_const(&ns_parts[0]);
                                    self.emit_u16(Op::GLOBAL_GET, root_idx);
                                    for part in &ns_parts[1..ns_parts.len() - 1] {
                                        let idx = self.str_const(part);
                                        self.emit_u16(Op::STRUCT_GET, idx);
                                    }
                                    let method_idx = self.str_const(&method_name);
                                    inst!(self, core_wasm::dup);
                                    self.emit_u16(Op::STRUCT_GET, method_idx);
                                    let fn_tmp = self.define_local("__ns_fn");
                                    self.emit_u16(Op::LOCAL_SET, fn_tmp);
                                    let obj_tmp = self.define_local("__ns_obj");
                                    self.reserve_local_slot(obj_tmp);
                                    self.emit_u16(Op::LOCAL_SET, obj_tmp);
                                    self.emit_u16(Op::LOCAL_GET, fn_tmp);
                                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                                    for a in &arg_exprs {
                                        self.compile_expr(a)?;
                                    }
                                    self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                                    return Ok(());
                                }

                                let root_idx = self.str_const(&ns_parts[0]);
                                self.emit_u16(Op::GLOBAL_GET, root_idx);
                                for part in &ns_parts[1..] {
                                    let idx = self.str_const(part);
                                    self.emit_u16(Op::STRUCT_GET, idx);
                                }
                                let is_const = ns_parts
                                    .last()
                                    .map(|name| dotnet_surface.is_known_constant(name))
                                    .unwrap_or(false);
                                if !is_const {
                                    for a in &arg_exprs {
                                        self.compile_expr(a)?;
                                    }
                                    self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                                }
                                return Ok(());
                            }
                            common::dotnet::DottedResolution::InstanceMember { local, members } => {
                                // Intercept `parent.Controls.Add(child)` for GUI.
                                // The .NET WinForms surface is `Form.Controls.Add(ctrl)`,
                                // MAUI is `parent.Children.Add(ctrl)`, etc. — all
                                // resolve to the canonical gui emitter.
                                if members.len() >= 2
                                    && members[members.len() - 2] == "controls"
                                    && members[members.len() - 1] == "add"
                                {
                                    let line = self.line;
                                    let add_idx =
                                        self.import("vybe:gui", common::gui::HOST_FN_ADD_CHILD);
                                    self.emit_var_get(&local);
                                    for a in &arg_exprs {
                                        self.compile_expr(a)?;
                                    }
                                    common::gui::emit_add_child(self.chunk(), add_idx, line);
                                    return Ok(());
                                }
                                // Intercept Thread/Task methods → WASM stack switching opcodes.
                                // Disambiguation by arity: `Thread.Join()` is zero-arg; an
                                // array's `.join(sep)` takes one. Without the arity gate
                                // this branch greedy-matched both and routed string-join
                                // through `thread.join` (which returns the exit code, not
                                // a string).
                                if members.len() == 1 && arg_exprs.is_empty() {
                                    let method = self.canon(members[0].as_str());
                                    match method.as_str() {
                                        "start" => {
                                            self.emit_var_get(&local);
                                            let line = self.line;
                                            common::threading::emit_thread_start(
                                                self.chunk(),
                                                line,
                                            );
                                            return Ok(());
                                        }
                                        "join" => {
                                            self.emit_var_get(&local);
                                            let line = self.line;
                                            common::threading::emit_thread_join(self.chunk(), line);
                                            return Ok(());
                                        }
                                        "waitforexit" => {
                                            self.emit_var_get(&local);
                                            let line = self.line;
                                            common::dotnet::core::process_adapter::emit_process_wait_for_exit(&mut self.chunks, self.current, line);
                                            return Ok(());
                                        }
                                        _ => {}
                                    }
                                }
                                let _ = local;
                                let _ = members;
                                // For ordinary local/member calls, fall through to the
                                // shared call pipeline below. That keeps value-method
                                // dispatch (`dict.Add`, `queue.Dequeue`, etc.) and the
                                // generic object member path as the single source of truth.
                            }
                            common::dotnet::DottedResolution::NoOp => {
                                self.emit(Op::NULL);
                                return Ok(());
                            }
                            common::dotnet::DottedResolution::Unresolved => {
                                // Fall through to value methods and other resolution
                            }
                        }
                    }
                }

                // Non-dotnet: namespace aliases (JS: console → wasi:cli).
                // Reads from `host_namespace_aliases` (populated by the
                // Linker) instead of `profile.lookup_module_alias` — one
                // source of truth for Member-chain resolution.
                let dotnet_root = self.profile.namespaces.use_dotnet_resolver
                    && common::dotnet::is_namespace_root(&lower_parts[0]);
                if !dotnet_root {
                    let alias_key = self.canon(&lower_parts[0]);
                    if let Some(module) = self.host_namespace_aliases.get(&alias_key).cloned() {
                        let is_js_prototype_chain = self.is_js_profile()
                            && lower_parts.len() > 2
                            && lower_parts
                                .get(1)
                                .is_some_and(|part| part.eq_ignore_ascii_case("prototype"));
                        let is_js_function_helper_chain = self.is_js_profile()
                            && lower_parts.len() > 2
                            && lower_parts.last().is_some_and(|part| {
                                matches!(part.as_str(), "call" | "apply" | "bind")
                            });
                        if is_js_prototype_chain || is_js_function_helper_chain {
                            // `Array.prototype.join.bind(...)` and similar borrowed-method
                            // chains must stay as property access on the extracted function
                            // value, not collapse into a synthetic host import like
                            // `ecma:array.prototype.join.bind` or `ecma:math.max.apply`.
                        } else {
                            // Check if any prefix of parts[0..n] is a namespace constant
                            // e.g. Math.PI.toFixed(5) — "Math.PI" is constant 3.14159, "toFixed" is method.
                            if lower_parts.len() > 2 {
                                let mut handled = false;
                                for end in (2..lower_parts.len()).rev() {
                                    let const_key = parts[..end].join(".");
                                    if let Some(cv) =
                                        self.profile.lookup_constant(&const_key).cloned()
                                    {
                                        match &cv {
                                            ConstantValue::Float(f) => {
                                                self.emit_const(Value::F64(*f))
                                            }
                                            ConstantValue::Str(s) => self
                                                .emit_const(Value::String(Arc::from(s.as_str()))),
                                        }
                                        let method_name = &parts[end];
                                        let argc = arg_exprs.len() as u8;
                                        let def = self
                                            .profile
                                            .lookup_value_method(method_name, argc)
                                            .cloned();
                                        if let Some(def) = def {
                                            for a in &arg_exprs {
                                                self.compile_expr(a)?;
                                            }
                                            match &def.emit {
                                                BuiltinEmit::HostCall(hmod, hfunc) => {
                                                    let (hmod, hfunc) =
                                                        (hmod.clone(), hfunc.clone());
                                                    let idx = self.import(&hmod, &hfunc);
                                                    self.emit_host_call(
                                                        idx,
                                                        (arg_exprs.len() + 1) as u8,
                                                    );
                                                }
                                                _ => {
                                                    let midx = self.str_const(method_name);
                                                    self.emit_u16(Op::STRUCT_GET, midx);
                                                    self.emit_u8(Op::CALL_REF, argc);
                                                }
                                            }
                                        } else {
                                            let midx = self.str_const(method_name);
                                            self.emit_u16(Op::STRUCT_GET, midx);
                                            for a in &arg_exprs {
                                                self.compile_expr(a)?;
                                            }
                                            self.emit_u8(Op::CALL_REF, argc);
                                        }
                                        handled = true;
                                        break;
                                    }
                                }
                                if handled {
                                    return Ok(());
                                }
                            }
                            let func = if lower_parts.len() == 2 {
                                lower_parts[1].clone()
                            } else {
                                lower_parts[1..].join(".")
                            };
                            let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                            for (index, arg) in arg_exprs.iter().enumerate() {
                                self.compile_expr(arg)?;
                                let arg_slot =
                                    self.define_local(&format!("__host_alias_arg_{}", index));
                                self.emit_u16(Op::LOCAL_SET, arg_slot);
                                arg_slots.push(arg_slot);
                            }
                            for slot in &arg_slots {
                                self.emit_u16(Op::LOCAL_GET, *slot);
                            }
                            let idx = self.import(&module, &func);
                            self.emit_host_call(idx, arg_exprs.len() as u8);
                            return Ok(());
                        }
                    }
                }

                // Profile namespace roots
                if self.profile.is_namespace_root(&lower_parts[0]) {
                    let root_idx = self.str_const(&lower_parts[0]);
                    self.emit_u16(Op::GLOBAL_GET, root_idx);
                    for part in &lower_parts[1..] {
                        let idx = self.str_const(part);
                        self.emit_u16(Op::STRUCT_GET, idx);
                    }
                    for a in &arg_exprs {
                        self.compile_expr(a)?;
                    }
                    self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                    return Ok(());
                }
            }
        }

        // ── Static method call on user class: ClassName.Method(args) ─
        // Must run BEFORE value methods so user class names like MathUtils.Add
        // don't get hijacked by the array Add value method.
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            let class_parts = self.flatten_member_chain(object);
            if !class_parts.is_empty() {
                let class_path = class_parts.join(".");
                let mut static_class_canon = None;
                let head_name = class_parts.first().map(String::as_str).unwrap_or("");
                // If any part of the chain (after the head) is a private field, this is
                // ClassName.#privateField.method(...) — the receiver is the private field
                // value, NOT the class itself. Don't treat it as a static method call.
                let chain_through_private = class_parts.iter().skip(1).any(|p| p.starts_with('#'));

                if !chain_through_private {
                    if let Some(current_class) = self.current_class.clone() {
                        if self.canon(head_name) == self.canon(&current_class)
                            || class_parts
                                .last()
                                .is_some_and(|part| self.canon(part) == self.canon(&current_class))
                        {
                            static_class_canon = Some(current_class);
                        }
                    }

                    let full_canon = self.canon(&class_path);
                    if static_class_canon.is_none()
                        && (self.defined_classes.contains(&full_canon)
                            || self.pending_classes.contains_key(&full_canon))
                        && self.scope().resolve(head_name).is_none()
                        && self.scope().resolve_ci(head_name).is_none()
                        && self.lookup_var_type_hint(head_name).is_none()
                    {
                        static_class_canon = Some(full_canon);
                    }

                    if static_class_canon.is_none() && class_parts.len() > 1 {
                        let short_name = class_parts.last().map(String::as_str).unwrap_or("");
                        let short_canon = self.canon(short_name);
                        if (self.defined_classes.contains(&short_canon)
                            || self.pending_classes.contains_key(&short_canon))
                            && self.scope().resolve(short_name).is_none()
                            && self.scope().resolve_ci(short_name).is_none()
                            && self.lookup_var_type_hint(short_name).is_none()
                        {
                            static_class_canon = Some(short_canon);
                        }
                    }
                }

                // §20.2.3 (JS): bind/call/apply on a class object resolve
                // through %Function.prototype% unless the class defines an
                // own static with that name — clear the static-dispatch
                // route so the generic Function.prototype path below
                // handles them (class constructors are function objects).
                if self.is_js_profile() && matches!(field.as_str(), "bind" | "call" | "apply") {
                    if let Some(canon) = static_class_canon.as_ref() {
                        let canon_field = self.canon(field);
                        let shadowed = self
                            .pending_classes
                            .get(canon.as_str())
                            .is_some_and(|pc| {
                                pc.static_method_overloads.contains_key(&canon_field)
                            });
                        if !shadowed {
                            static_class_canon = None;
                        }
                    }
                }

                if let Some(canon) = static_class_canon {
                    if self.is_js_profile() {
                        let method_name = self.js_member_storage_name_for_class(&canon, field);
                        let cls_idx = self.str_const(&canon);
                        self.emit_u16(Op::GLOBAL_GET, cls_idx);
                        let cls_tmp = self
                            .scope()
                            .resolve("__static_cls")
                            .unwrap_or_else(|| self.define_local("__static_cls"));
                        self.emit_u16(Op::LOCAL_SET, cls_tmp);
                        let fn_tmp = self
                            .scope()
                            .resolve("__static_fn")
                            .unwrap_or_else(|| self.define_local("__static_fn"));
                        if field.starts_with('#') {
                            if let Some(overload) = self
                                .resolve_static_method_overload_for_type(&canon, field, &arg_exprs)
                            {
                                let line = self.line;
                                self.emit_u16(Op::REF_FUNC, overload.chunk_idx as u16);
                                self.chunk().emit(0, line);
                            } else {
                                self.emit_u16(Op::LOCAL_GET, cls_tmp);
                                let method_idx = self.str_const(&method_name);
                                self.emit_u16(Op::STRUCT_GET, method_idx);
                            }
                        } else {
                            self.emit_u16(Op::LOCAL_GET, cls_tmp);
                            let method_idx = self.str_const(&method_name);
                            self.emit_u16(Op::STRUCT_GET, method_idx);
                        }
                        self.emit_u16(Op::LOCAL_SET, fn_tmp);
                        let saved_js_this = self.save_js_this("__js_prev_this_static_method");
                        self.emit_u16(Op::LOCAL_GET, cls_tmp);
                        self.set_js_this_from_stack();
                        let qualified_method = self.canon(&format!("{}.{}", canon, field));
                        if let Some(param_modes) = self
                            .function_param_modes
                            .get(&qualified_method)
                            .cloned()
                            .or_else(|| self.function_param_modes.get(&method_name).cloned())
                        {
                            if param_modes
                                .iter()
                                .any(|mode| matches!(mode, PassBy::Ref | PassBy::Out))
                            {
                                let mut arg_slots = Vec::with_capacity(args.len());
                                for (index, arg) in args.iter().enumerate() {
                                    self.compile_ref_aware_call_arg(
                                        arg,
                                        param_modes.get(index).copied().unwrap_or(PassBy::Value),
                                    )?;
                                    let arg_slot = self
                                        .define_local(&format!("__js_static_call_arg_{}", index));
                                    self.emit_u16(Op::LOCAL_SET, arg_slot);
                                    arg_slots.push(arg_slot);
                                }

                                self.emit_u16(Op::LOCAL_GET, fn_tmp);
                                for slot in &arg_slots {
                                    self.emit_u16(Op::LOCAL_GET, *slot);
                                }
                                self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                                let pack_slot = self.define_local("__js_static_ref_call_pack");
                                self.emit_u16(Op::LOCAL_SET, pack_slot);
                                self.restore_js_this(saved_js_this);

                                let mut ref_out_index = 1usize;
                                for (index, arg) in args.iter().enumerate() {
                                    if !matches!(
                                        param_modes.get(index),
                                        Some(PassBy::Ref | PassBy::Out)
                                    ) {
                                        continue;
                                    }
                                    self.emit_u16(Op::LOCAL_GET, pack_slot);
                                    self.emit_const(Value::F64(ref_out_index as f64));
                                    common::collections::emit_get(
                                        &mut self.chunks,
                                        self.current,
                                        self.line,
                                    );
                                    self.compile_assign_target(&arg.value)?;
                                    ref_out_index += 1;
                                }

                                self.emit_u16(Op::LOCAL_GET, pack_slot);
                                self.emit_const(Value::F64(0.0));
                                common::collections::emit_get(
                                    &mut self.chunks,
                                    self.current,
                                    self.line,
                                );
                                return Ok(());
                            }
                        }

                        self.emit_u16(Op::LOCAL_GET, fn_tmp);
                        for a in &arg_exprs {
                            self.compile_expr(a)?;
                        }
                        self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                        let result_slot = self.define_local("__js_static_method_result");
                        self.emit_u16(Op::LOCAL_SET, result_slot);
                        self.restore_js_this(saved_js_this);
                        if args.iter().any(|arg| arg.by_ref) {
                            let mut ref_out_index = 1usize;
                            for arg in args {
                                if !arg.by_ref {
                                    continue;
                                }
                                self.emit_u16(Op::LOCAL_GET, result_slot);
                                self.emit_const(Value::F64(ref_out_index as f64));
                                common::collections::emit_get(
                                    &mut self.chunks,
                                    self.current,
                                    self.line,
                                );
                                self.compile_assign_target(&arg.value)?;
                                ref_out_index += 1;
                            }
                            self.emit_u16(Op::LOCAL_GET, result_slot);
                            self.emit_const(Value::F64(0.0));
                            common::collections::emit_get(
                                &mut self.chunks,
                                self.current,
                                self.line,
                            );
                            return Ok(());
                        }
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                        return Ok(());
                    }

                    let cls_idx = self.str_const(&canon);
                    self.emit_u16(Op::GLOBAL_GET, cls_idx);
                    inst!(self, core_wasm::dup);
                    let m = self.canon(field);
                    let method_idx = self.str_const(&m);
                    self.emit_u16(Op::STRUCT_GET, method_idx);
                    // Stack: [class, fn] — swap so we have [fn, class, ...args]
                    let fn_tmp = self
                        .scope()
                        .resolve("__static_fn")
                        .unwrap_or_else(|| self.define_local("__static_fn"));
                    self.emit_u16(Op::LOCAL_SET, fn_tmp);
                    let cls_tmp = self
                        .scope()
                        .resolve("__static_cls")
                        .unwrap_or_else(|| self.define_local("__static_cls"));
                    self.emit_u16(Op::LOCAL_SET, cls_tmp);
                    let qualified_method = self.canon(&format!("{}.{}", canon, field));
                    if let Some(param_modes) = self
                        .function_param_modes
                        .get(&qualified_method)
                        .cloned()
                        .or_else(|| self.function_param_modes.get(&m).cloned())
                    {
                        if param_modes
                            .iter()
                            .any(|mode| matches!(mode, PassBy::Ref | PassBy::Out))
                        {
                            let mut arg_slots = Vec::with_capacity(args.len());
                            for (index, arg) in args.iter().enumerate() {
                                self.compile_ref_aware_call_arg(
                                    arg,
                                    param_modes.get(index).copied().unwrap_or(PassBy::Value),
                                )?;
                                let arg_slot =
                                    self.define_local(&format!("__static_call_arg_{}", index));
                                self.emit_u16(Op::LOCAL_SET, arg_slot);
                                arg_slots.push(arg_slot);
                            }

                            self.emit_u16(Op::LOCAL_GET, fn_tmp);
                            for slot in &arg_slots {
                                self.emit_u16(Op::LOCAL_GET, *slot);
                            }
                            self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);

                            let pack_slot = self.define_local("__static_ref_call_pack");
                            self.emit_u16(Op::LOCAL_SET, pack_slot);
                            let mut ref_out_index = 1usize;
                            for (index, arg) in args.iter().enumerate() {
                                if !matches!(
                                    param_modes.get(index),
                                    Some(PassBy::Ref | PassBy::Out)
                                ) {
                                    continue;
                                }
                                self.emit_u16(Op::LOCAL_GET, pack_slot);
                                self.emit_const(Value::F64(ref_out_index as f64));
                                common::collections::emit_get(
                                    &mut self.chunks,
                                    self.current,
                                    self.line,
                                );
                                self.compile_assign_target(&arg.value)?;
                                ref_out_index += 1;
                            }
                            self.emit_u16(Op::LOCAL_GET, pack_slot);
                            self.emit_const(Value::F64(0.0));
                            common::collections::emit_get(
                                &mut self.chunks,
                                self.current,
                                self.line,
                            );
                            return Ok(());
                        }
                    }

                    if self.profile.name == "csharp" && args.len() == 1 && !args[0].spread {
                        if self
                            .resolve_static_method_overload_for_type(&canon, field, &arg_exprs)
                            .is_some_and(|overload| overload.signature.has_rest)
                        {
                            self.emit_variadic_array_call_from_local(fn_tmp, &args[0].value)?;
                            return Ok(());
                        }
                    }
                    let rest_signature = self
                        .function_signatures
                        .get(&qualified_method)
                        .and_then(|signatures| self.select_call_signature(signatures, args))
                        .filter(|signature| signature.has_rest)
                        .cloned()
                        .or_else(|| {
                            self.function_signatures
                                .get(&m)
                                .and_then(|signatures| self.select_call_signature(signatures, args))
                                .filter(|signature| signature.has_rest)
                                .cloned()
                        });
                    if let Some(signature) = rest_signature.as_ref() {
                        self.emit_known_rest_call_from_local(fn_tmp, None, args, signature)?;
                    } else {
                        let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                        for (index, arg) in arg_exprs.iter().enumerate() {
                            self.compile_expr(arg)?;
                            let arg_slot =
                                self.define_local(&format!("__static_class_arg_{}", index));
                            self.emit_u16(Op::LOCAL_SET, arg_slot);
                            arg_slots.push(arg_slot);
                        }
                        self.emit_call_ref_with_arg_slots(fn_tmp, None, &arg_slots);
                    }
                    if args.iter().any(|arg| arg.by_ref) {
                        let pack_slot = self.define_local("__static_by_ref_pack");
                        self.emit_u16(Op::LOCAL_SET, pack_slot);
                        let mut ref_out_index = 1usize;
                        for arg in args {
                            if !arg.by_ref {
                                continue;
                            }
                            self.emit_u16(Op::LOCAL_GET, pack_slot);
                            self.emit_const(Value::F64(ref_out_index as f64));
                            common::collections::emit_get(
                                &mut self.chunks,
                                self.current,
                                self.line,
                            );
                            self.compile_assign_target(&arg.value)?;
                            ref_out_index += 1;
                        }
                        self.emit_u16(Op::LOCAL_GET, pack_slot);
                        self.emit_const(Value::F64(0.0));
                        common::collections::emit_get(&mut self.chunks, self.current, self.line);
                    }
                    return Ok(());
                }
            }
        }

        // ── Nested static type call: Outer.Inner.Method(args) ───────
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            if let ExprKind::Member {
                object: outer_obj,
                field: nested_name,
                ..
            } = &object.kind
            {
                if let ExprKind::Ident(outer_name) = &outer_obj.kind {
                    let outer_canon = self.canon(outer_name);
                    let is_outer_class = self.defined_classes.contains(&outer_canon)
                        && self.scope().resolve(outer_name).is_none();
                    if is_outer_class {
                        let nested_ok = self
                            .pending_classes
                            .get(outer_canon.as_str())
                            .map(|pc| {
                                pc.nested_types.iter().any(|n| {
                                    if self.case_sensitive {
                                        n == nested_name
                                    } else {
                                        n.eq_ignore_ascii_case(nested_name)
                                    }
                                })
                            })
                            .unwrap_or(false);
                        if nested_ok {
                            let outer_idx = self.str_const(&outer_canon);
                            self.emit_u16(Op::GLOBAL_GET, outer_idx);
                            let nested_idx = self.str_const(&self.canon(nested_name));
                            self.emit_u16(Op::STRUCT_GET, nested_idx);
                            let cls_tmp = self
                                .scope()
                                .resolve("__nested_static_cls")
                                .unwrap_or_else(|| self.define_local("__nested_static_cls"));
                            self.emit_u16(Op::LOCAL_SET, cls_tmp);
                            self.emit_u16(Op::LOCAL_GET, cls_tmp);
                            let method_idx = self.str_const(&self.canon(field));
                            self.emit_u16(Op::STRUCT_GET, method_idx);
                            let fn_tmp = self
                                .scope()
                                .resolve("__nested_static_fn")
                                .unwrap_or_else(|| self.define_local("__nested_static_fn"));
                            self.emit_u16(Op::LOCAL_SET, fn_tmp);
                            self.emit_u16(Op::LOCAL_GET, fn_tmp);
                            for a in &arg_exprs {
                                self.compile_expr(a)?;
                            }
                            self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                            return Ok(());
                        }
                    }
                }
            }
        }

        // ── Function.prototype.call / .apply ────────────────────────
        // `fn.call(thisArg, a, b, ...)` → call `fn` with `[a, b, ...]`
        // `fn.apply(thisArg, [a, b, ...])` → same; the array form is
        // unwrapped at runtime via the spread opcode.
        //
        // We can't route this through value_methods because the standard
        // dispatch path pushes the receiver + ALL args, but here we need
        // to drop arg[0] (`thisArg`) from the middle of the stack. Skip
        // when the field is defined on a user class so user methods
        // named `call`/`apply` keep working.
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            if self.is_js_profile()
                && !self.direct_receiver_has_own_pending_method(object, field)
                && (field == "call" || field == "apply" || field == "bind")
            {
                // §20.2.3.{1,2,3}: call/apply/bind live on
                // %Function.prototype%. Function objects don't all carry a
                // proto link to it (static / prototype methods are bare
                // REF_FUNC values), so route the call form directly to the
                // host — `ecma:function.<m>(target, ...)`.
                self.compile_expr(object)?;
                for arg in &arg_exprs {
                    self.compile_expr(arg)?;
                }
                let idx = self.import("ecma:function", field);
                self.emit_host_call(idx, (arg_exprs.len() + 1) as u8);
                return Ok(());
            }
            if !self.direct_receiver_has_own_pending_method(object, field)
                && (field == "call" || field == "apply")
            {
                self.compile_expr(object)?;
                for arg in &arg_exprs {
                    self.compile_expr(arg)?;
                }
                let idx = self.import("ecma:function", field);
                self.emit_host_call(idx, (arg_exprs.len() + 1) as u8);
                return Ok(());
            }
        }

        // ── Component Model instance-method dispatch ────────────────
        //
        // When `obj` is a local with a known .NET type (from
        // `Dim d As New Dictionary(...)` / `var x : Stack` / etc.),
        // resolve the method against the auto-built component
        // descriptor and emit the import call directly. This is the
        // primary dispatch path per the Component Model + ESM
        // architecture — the .NET adapter at the descriptor level
        // translates `Dictionary.Add` → `ecma:map.set`, so the
        // emitted call hits the standardized primitive without any
        // runtime `__type` lookup. The TypeRegistry-driven runtime
        // dispatch (compilation-hints proposal style) is the
        // fallback for dynamically-typed receivers.
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            let class_name = resolve_receiver_type_hint(self, object);
            if let Some(class_name) = class_name {
                if self
                    .resolve_pending_class_name_for_type_hint(&class_name)
                    .is_some()
                {
                    // User-defined classes win over shared .NET surface names
                    // like `Stack`, `Queue`, or `Dictionary`.
                } else {
                    let class_name = Self::normalize_type_hint(&class_name);
                    let surface = common::dotnet::surface();
                    if let Some(target) =
                        surface.lookup_instance_method(&class_name, field, arg_exprs.len() as u8)
                    {
                        if matches!(&target, common::dotnet::InstanceMethodTarget::Common { emit, .. } if emit == "collections.sort")
                            && arg_exprs.is_empty()
                            && !self.is_js_profile()
                        {
                            let sort_global = self.str_const("__vybe_sort_with_comparator");
                            self.emit_u16(Op::GLOBAL_GET, sort_global);
                            self.compile_expr(object)?;
                            self.compile_lambda(
                                &[
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
                                &LambdaBody::Expr(Box::new(Expression::new(ExprKind::Ternary {
                                    cond: Box::new(Expression::new(ExprKind::Binary {
                                        op: BinOp::Lt,
                                        left: Box::new(Expression::ident("left")),
                                        right: Box::new(Expression::ident("right")),
                                    })),
                                    then: Box::new(Expression::new(ExprKind::Lit(Literal::Int(
                                        -1,
                                    )))),
                                    else_: Box::new(Expression::new(ExprKind::Ternary {
                                        cond: Box::new(Expression::new(ExprKind::Binary {
                                            op: BinOp::Gt,
                                            left: Box::new(Expression::ident("left")),
                                            right: Box::new(Expression::ident("right")),
                                        })),
                                        then: Box::new(Expression::new(ExprKind::Lit(
                                            Literal::Int(1),
                                        ))),
                                        else_: Box::new(Expression::new(ExprKind::Lit(
                                            Literal::Int(0),
                                        ))),
                                    })),
                                }))),
                                &[],
                            )?;
                            self.emit_u8(Op::CALL_REF, 2);
                            return Ok(());
                        }

                        if matches!(&target, common::dotnet::InstanceMethodTarget::Common { emit, .. } if emit == "dotnet.array_sort")
                            && arg_exprs.len() == 1
                            && !self.is_js_profile()
                            && class_name.rsplit('.').next().is_some_and(|name| {
                                name.eq_ignore_ascii_case("List")
                                    || name.eq_ignore_ascii_case("ArrayList")
                            })
                            && matches!(&arg_exprs[0].kind, ExprKind::Lambda { params, .. } if params.len() == 2)
                        {
                            let sort_global = self.str_const("__vybe_sort_with_comparator");
                            self.emit_u16(Op::GLOBAL_GET, sort_global);
                            self.compile_expr(object)?;
                            self.compile_expr(&arg_exprs[0])?;
                            self.emit_u8(Op::CALL_REF, 2);
                            return Ok(());
                        }

                        // Compile receiver, then args.
                        self.compile_expr(object)?;
                        for a in &arg_exprs {
                            self.compile_expr(a)?;
                        }
                        let total_argc = (arg_exprs.len() + 1) as u8;
                        match target {
                            common::dotnet::InstanceMethodTarget::Host { module, func, .. } => {
                                let idx = self.import(&module, &func);
                                self.emit_host_call(idx, total_argc);
                            }
                            common::dotnet::InstanceMethodTarget::Common { emit, .. } => {
                                let line = self.line;
                                self.emit_common(&emit, total_argc, line);
                            }
                        }
                        return Ok(());
                    }
                }
            }
        }

        // ── Value method: obj.toUpperCase() ─────────────────────────
        //
        // Method name shadowing rule: a value method (e.g. `Array.push`,
        // `String.toUpperCase`) is the default for *member-access*
        // receivers like `this.items.push(x)` — the receiver is
        // structurally a property, almost certainly a built-in collection.
        //
        // For *direct* receivers (`this`, `super`, or a local variable
        // by name), if the field is a known user-class method, prefer
        // the user method via the generic call path. That preserves
        // user overrides like `class Stack { push(x) { ... } }` and
        // `class Holder { size() { ... } }` against built-in
        // `Array.push`/`map_size` shadowing.
        //
        // This is a heuristic — the cleaner fix is per-class method sets
        // plus receiver-type inference, tracked in the user's pending
        // "JS/C# compilers don't use common::classes" migration.
        if let ExprKind::Member {
            object,
            field,
            null_safe,
        } = &callee.kind
        {
            let canon_field = self.canon(field);
            let receiver_is_direct = matches!(
                object.kind,
                ExprKind::This | ExprKind::Super | ExprKind::Ident(_)
            );
            if self.is_python_profile() && arg_exprs.is_empty() {
                if let ExprKind::Lit(Literal::Str(value)) = &object.kind {
                    match field.as_str() {
                        "isidentifier" => {
                            self.emit_const(Value::Bool(python_is_identifier_literal(
                                value.as_ref(),
                            )));
                            return Ok(());
                        }
                        "isprintable" => {
                            self.emit_const(Value::Bool(python_is_printable_literal(
                                value.as_ref(),
                            )));
                            return Ok(());
                        }
                        _ => {}
                    }
                }
            }
            // Skip value-method dispatch on null-safe member calls — the
            // null short-circuit must run BEFORE we apply any built-in
            // operator (e.g. `null?.toUpperCase()` returns null, not "").
            // Falls through to the generic Member-access path which
            // handles null_safe correctly.
            let matched_value_method = if *null_safe {
                None
            } else {
                self.profile
                    .lookup_value_method(field, arg_exprs.len() as u8)
                    .cloned()
            };
            let array_only_value_method_for_non_array = matches!(
                matched_value_method.as_ref().map(|d| &d.emit),
                Some(BuiltinEmit::HostCall(module, func))
                    if module == "ecma:array"
                        && func == "entries"
                        && !self.expr_is_array_like(object)
            );
            // Keep dotnet adapter value-methods ahead of runtime collection
            // dispatch for untyped receivers (notably plain arrays using
            // LINQ-style extension methods like Select/SelectMany).
            let prefer_dotnet_adapter = match matched_value_method.as_ref().map(|d| &d.emit) {
                Some(BuiltinEmit::Common(name)) => name.starts_with("dotnet."),
                _ => false,
            };
            let receiver_type_hint = self.infer_expr_type_hint(object);
            let receiver_has_pending_user_method = self
                .infer_expr_type_hint(object)
                .as_deref()
                .is_some_and(|type_hint| {
                    self.pending_class_has_method_name_for_type(type_hint, field)
                });
            let receiver_is_user_type = self
                .infer_expr_type_hint(object)
                .as_deref()
                .and_then(|type_hint| self.resolve_pending_class_name_for_type_hint(type_hint))
                .is_some();
            let receiver_is_known_builtin_value = receiver_type_hint
                .as_deref()
                .map(Self::normalize_type_hint)
                .is_some_and(|type_hint| {
                    Self::is_collection_like_type_hint(&type_hint)
                        || Self::is_string_type_hint(&type_hint)
                        || matches!(
                            type_hint.as_str(),
                            "number" | "int" | "double" | "bool" | "boolean"
                        )
                });
            let receiver_is_url_search_params = receiver_type_hint
                .as_deref()
                .map(Self::normalize_type_hint)
                .is_some_and(|type_hint| type_hint == "urlsearchparams")
                || matches!(&object.kind,
                    ExprKind::Member { object: url_object, field: member_field, .. }
                        if member_field == "searchParams"
                            && self
                                .infer_expr_type_hint(url_object)
                                .as_deref()
                                .map(Self::normalize_type_hint)
                                .is_some_and(|type_hint| type_hint == "url")
                );
            let user_method_shadow = self.direct_receiver_has_own_pending_method(object, field)
                || receiver_has_pending_user_method
                || (receiver_is_direct
                    && !receiver_is_known_builtin_value
                    && self.defined_class_methods.contains(&canon_field))
                || (receiver_is_direct
                    && receiver_is_user_type
                    && self.defined_class_methods.contains(&canon_field));
            // Also skip value_methods if the field is an array HOF method —
            // the array_methods dispatch handles it with proper HOF semantics.
            // Without this, `[1,2,3].includes(2)` routes through the string
            // `includes` value method instead of the array contains HOF.
            let field_lower_check = if self.case_sensitive {
                field.clone()
            } else {
                field.to_lowercase()
            };
            let is_array_method = self
                .profile
                .lookup_array_method(&field_lower_check)
                .is_some();
            if user_method_shadow || is_array_method {
                // Fall through — let the HOF dispatch or generic call path handle it
            } else if array_only_value_method_for_non_array {
                // Array-only value methods like `.entries()` must not steal
                // Map/Set receivers away from runtime method dispatch.
            } else if self.profile.namespaces.use_dotnet
                && common::dotnet::uses_runtime_collection_dispatch_arity(
                    field,
                    arg_exprs.len() as u8,
                )
                && !prefer_dotnet_adapter
            {
                // Let the generic member-call path consult the runtime type
                // registry for shared .NET collection methods instead of
                // intercepting them via language profile value-method tables.
            } else if let Some(def) = matched_value_method {
                if self.is_js_profile() && field == "push" && args.iter().any(|arg| arg.spread) {
                    let line = self.line;
                    self.compile_expr(object)?;
                    let obj_slot = self.define_local("__js_value_push_spread_obj");
                    self.emit_u16(Op::LOCAL_SET, obj_slot);

                    let (args_slot, _) =
                        self.compile_call_args_array(args, "js_value_push_spread_values")?;
                    let len_slot = self.define_local("__js_value_push_spread_len");
                    self.emit_u16(Op::LOCAL_GET, args_slot);
                    common::collections::emit_len(&mut self.chunks, self.current, line);
                    self.emit_u16(Op::LOCAL_SET, len_slot);

                    let idx_slot = self.define_local("__js_value_push_spread_idx");
                    self.emit_const(Value::I32(0));
                    self.emit_u16(Op::LOCAL_SET, idx_slot);

                    let loop_block = self.chunk().emit_block(line);
                    let (loop_patch, _) = self.chunk().emit_loop_s(line);
                    self.emit_u16(Op::LOCAL_GET, idx_slot);
                    self.emit_u16(Op::LOCAL_GET, len_slot);
                    {
                        let line = self.line;
                        crate::emitter::ops::emit_dyn_lt(self.chunk(), line);
                    };
                    let line = self.line;
                    crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                    crate::emitter::ops::emit_dyn_not(self.chunk(), line);
                    self.chunk().emit_br_if(1, line);

                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit_u16(Op::LOCAL_GET, args_slot);
                    self.emit_u16(Op::LOCAL_GET, idx_slot);
                    common::collections::emit_get(&mut self.chunks, self.current, line);
                    common::collections::emit_push(&mut self.chunks, self.current, line);
                    self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, idx_slot);
                    self.emit_const(Value::I32(1));
                    {
                        let line = self.line;
                        crate::emitter::ops::emit_dyn_add(self.chunk(), line);
                    };
                    self.emit_u16(Op::LOCAL_SET, idx_slot);
                    self.chunk().emit_br(0, line);
                    self.chunk().emit_end(line);
                    self.chunk().patch_loop(loop_patch);
                    self.chunk().emit_end(line);
                    self.chunk().patch_block(loop_block);

                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    common::collections::emit_len(&mut self.chunks, self.current, line);
                    return Ok(());
                }
                // Object is first arg, then explicit args
                self.compile_expr(object)?;
                for a in &arg_exprs {
                    self.compile_expr(a)?;
                }
                // Some opcodes need default args when called with fewer
                // than required. Push defaults here.
                if let BuiltinEmit::Opcode(op) | BuiltinEmit::Common(op) = &def.emit {
                    match op.as_str() {
                        // array_join / collections.join needs [arr, sep]
                        "array_join" | "collections.join" if arg_exprs.is_empty() => {
                            self.emit_const(Value::String(Arc::from(",")));
                        }
                        // array_fill needs [arr, val, start, end]
                        "array_fill" if arg_exprs.len() < 2 => {
                            // Push start=0 and end=arr.length defaults
                            if arg_exprs.is_empty() {
                                self.emit(Op::NULL); // val
                            }
                            inst!(self, core_wasm::i32_const, 0); // start
                            self.emit_const(Value::I32(i32::MAX)); // end (clamped by VM)
                        }
                        // C# `s.Substring(start)` — 1-arg form means
                        // "from start to end of string". STR_SUBSTRING
                        // wants `[s, start, end]`; default end to a
                        // sentinel large value (VM clamps to s.len()).
                        // Same shape applies to ECMA-262 §22.1.3.16
                        // `String.prototype.slice(start)`.
                        "strings.substring" | "strings.slice" if arg_exprs.len() < 2 => {
                            self.emit_const(Value::I32(i32::MAX));
                        }
                        // C#'s `string.ToCharArray()` lowers to STR_SPLIT
                        // which needs a delimiter on the stack. The .NET
                        // semantics ("each char one element") match
                        // splitting on the empty string.
                        "str_split" if arg_exprs.is_empty() => {
                            self.emit_const(Value::String(Arc::from("")));
                        }
                        _ => {}
                    }
                }
                match &def.emit {
                    BuiltinEmit::HostCall(module, func) => {
                        let idx = self.import(module, func);
                        self.emit_host_call(idx, (arg_exprs.len() + 1) as u8);
                    }
                    BuiltinEmit::Opcode(op_name) => {
                        // Object + args already on stack from above
                        self.emit_named_opcode(op_name);
                    }
                    BuiltinEmit::StrLength => {
                        let line = self.line;
                        common::strings::emit_length(self.chunk(), line);
                    }
                    BuiltinEmit::Common(name) => {
                        let line = self.line;
                        let name = name.clone();
                        self.emit_common(&name, (arg_exprs.len() + 1) as u8, line);
                    }
                    BuiltinEmit::Invoke(method_name) => {
                        let line = self.line;
                        let name = method_name.clone();
                        common::invoke::emit_invoke_method(
                            &mut self.chunks,
                            self.current,
                            &name,
                            arg_exprs.len() as u8,
                            line,
                        );
                    }
                    _ => {}
                }
                return Ok(());
            }

            // Array higher-order methods: arr.map(fn), arr.filter(fn), etc.
            // Use compiler_common::loops which emits proper loop bytecode.
            // BUT: skip when the same name is a user-defined class method
            // (e.g. `QueryBuilder.Where(string)` shouldn't be intercepted
            // by the LINQ HOF dispatch). The compiler can't see receiver
            // types at compile time, but it knows what method names user
            // classes have declared.
            let field_lower = if self.case_sensitive {
                field.clone()
            } else {
                field.to_lowercase()
            };
            let user_class_method = self.direct_receiver_has_own_pending_method(object, field)
                || self
                    .infer_expr_type_hint(object)
                    .as_deref()
                    .is_some_and(|type_hint| {
                        self.pending_class_has_method_name_for_type(type_hint, field)
                    });
            let js_requires_dynamic_callback_dispatch = self.is_js_profile()
                && arg_exprs
                    .first()
                    .is_some_and(|expr| matches!(expr.kind, ExprKind::Call { .. }));
            if !user_class_method
                && !receiver_is_url_search_params
                && !js_requires_dynamic_callback_dispatch
                && self.profile.lookup_array_method(&field_lower).is_some()
            {
                // (re-fetch only when we're committed to the HOF path so
                // the method name lookup matches the previous behaviour)
            }
            if let Some(stdlib_name) = self
                .profile
                .lookup_array_method(&field_lower)
                .filter(|_| {
                    !self.is_js_profile()
                        && !user_class_method
                        && !receiver_is_url_search_params
                        && !js_requires_dynamic_callback_dispatch
                })
                .map(|s| s.to_string())
            {
                // Normalize to the JS-style method name used in match below
                let field_lower = match stdlib_name.as_str() {
                    "__array_map" => "map".to_string(),
                    "__array_filter" => "filter".to_string(),
                    "__array_forEach" => "forEach".to_string(),
                    "__array_reduce" => "reduce".to_string(),
                    "__array_find" => "find".to_string(),
                    "__array_sort" => "sort".to_string(),
                    "__array_sort_by_key" => "sort_by_key".to_string(),
                    "__array_some" => "some".to_string(),
                    "__array_every" => "every".to_string(),
                    "__array_flat_map" => "flatMap".to_string(),
                    "__array_reduce_right" => "reduceRight".to_string(),
                    _ => field_lower,
                };
                // Compile arr and fn(s) into local slots
                self.compile_expr(object)?;
                let arr_slot = self.define_local("__hof_arr");
                self.emit_u16(Op::LOCAL_SET, arr_slot);

                if let Some(fn_expr) = arg_exprs.first() {
                    self.compile_expr(fn_expr)?;
                } else {
                    self.emit(Op::NULL);
                }
                let fn_slot = self.define_local("__hof_fn");
                self.emit_u16(Op::LOCAL_SET, fn_slot);

                let idx_slot = self.define_local("__hof_idx");
                let result_slot = self.define_local("__hof_result");
                let line = self.line;

                match field_lower.as_str() {
                    "map" => {
                        // emit_map leaves result on stack
                        common::loops::emit_map(
                            &mut self.chunks,
                            self.current,
                            fn_slot,
                            arr_slot,
                            result_slot,
                            idx_slot,
                            line,
                        );
                    }
                    "filter" => {
                        let elem_slot = self.define_local("__hof_elem");
                        common::loops::emit_filter(
                            &mut self.chunks,
                            self.current,
                            fn_slot,
                            arr_slot,
                            result_slot,
                            idx_slot,
                            elem_slot,
                            line,
                        );
                    }
                    "reduce" => {
                        // reduce(fn, initial?) — initial is second arg.
                        // When initial IS provided, start from i=0 with
                        // acc=initial. emit_reduce always starts from
                        // i=1 with acc=arr[0], so we only use it for
                        // the no-initial case.
                        if let Some(init_expr) = arg_exprs.get(1) {
                            // acc = initial, i = 0
                            self.compile_expr(init_expr)?;
                            self.emit_u16(Op::LOCAL_SET, result_slot);
                            // Inline reduce loop starting from i=0
                            inst!(self, core_wasm::i32_const, 0);
                            self.emit_u16(Op::LOCAL_SET, idx_slot);
                            let loop_block = self.chunk().emit_block(line);
                            let (loop_patch, _) = self.chunk().emit_loop_s(line);
                            self.emit_u16(Op::LOCAL_GET, idx_slot);
                            self.emit_u16(Op::LOCAL_GET, arr_slot);
                            {
                                let l = self.line;
                                common::collections::emit_len(&mut self.chunks, self.current, l);
                            }
                            {
                                let line = self.line;
                                crate::emitter::ops::emit_dyn_lt(self.chunk(), line);
                            };
                            let line = self.line;
                            crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                            crate::emitter::ops::emit_dyn_not(self.chunk(), line);
                            self.chunk().emit_br_if(1, line);
                            // acc = fn(acc, arr[i], i)  — ECMA-262 §23.1.3.26 passes (acc, elem, index, array)
                            self.emit_u16(Op::LOCAL_GET, fn_slot);
                            self.emit_u16(Op::LOCAL_GET, result_slot);
                            self.emit_u16(Op::LOCAL_GET, arr_slot);
                            self.emit_u16(Op::LOCAL_GET, idx_slot);
                            {
                                let l = self.line;
                                common::collections::emit_get(&mut self.chunks, self.current, l);
                            }
                            self.emit_u16(Op::LOCAL_GET, idx_slot);
                            self.emit_u8(Op::CALL_REF, 3);
                            self.emit_u16(Op::LOCAL_SET, result_slot);
                            // i++
                            self.emit_u16(Op::LOCAL_GET, idx_slot);
                            self.emit_const(Value::I32(1));
                            {
                                let line = self.line;
                                crate::emitter::ops::emit_dyn_add(self.chunk(), line);
                            };
                            self.emit_u16(Op::LOCAL_SET, idx_slot);
                            self.chunk().emit_br(0, line);
                            self.chunk().emit_end(line);
                            self.chunk().patch_loop(loop_patch);
                            self.chunk().emit_end(line);
                            self.chunk().patch_block(loop_block);
                            self.emit_u16(Op::LOCAL_GET, result_slot);
                        } else {
                            // No initial: emit_reduce starts from arr[0], i=1
                            common::loops::emit_reduce(
                                &mut self.chunks,
                                self.current,
                                fn_slot,
                                arr_slot,
                                result_slot,
                                idx_slot,
                                line,
                            );
                        }
                    }
                    "forEach" | "foreach" => {
                        // Polymorphic forEach: arrays iterate by index,
                        // Maps iterate (val, key, map) per ECMA-262
                        // §24.1.3.5, Sets iterate (val, val, set). The
                        // compiler can't know the receiver type so route
                        // through `ecma:value.invokeMethod` (each impl
                        // is in dispatch_{array,map,set}). For non-JS
                        // profiles, keep the array-only stdlib loop —
                        // PHP / VB iteration semantics differ.
                        if self.is_js_profile() {
                            self.emit_u16(Op::LOCAL_GET, arr_slot);
                            self.emit_u16(Op::LOCAL_GET, fn_slot);
                            if let Some(this_arg) = arg_exprs.get(1) {
                                self.compile_expr(this_arg)?;
                            }
                            common::invoke::emit_invoke_method(
                                &mut self.chunks,
                                self.current,
                                "forEach",
                                if arg_exprs.get(1).is_some() { 2 } else { 1 },
                                line,
                            );
                            self.emit(Op::DROP); // forEach returns undefined
                        } else {
                            common::loops::emit_foreach(
                                &mut self.chunks,
                                self.current,
                                fn_slot,
                                arr_slot,
                                idx_slot,
                                line,
                            );
                        }
                    }
                    "some" => {
                        common::loops::emit_any_every(
                            &mut self.chunks,
                            self.current,
                            fn_slot,
                            arr_slot,
                            idx_slot,
                            true,
                            line,
                        );
                    }
                    "every" => {
                        common::loops::emit_any_every(
                            &mut self.chunks,
                            self.current,
                            fn_slot,
                            arr_slot,
                            idx_slot,
                            false,
                            line,
                        );
                    }
                    "find" => {
                        // find uses includes pattern but returns element not bool.
                        // JS spec §23.1.3.10: returns undefined when no match;
                        // other languages stick with Null for cross-compat
                        // (Python None / VB Nothing / .NET null match Null).
                        if self.is_js_profile() {
                            inst!(self, core_wasm::undefined);
                        } else {
                            self.emit(Op::NULL);
                        }
                        self.emit_u16(Op::LOCAL_SET, result_slot);
                        let lp = common::loops::emit_for_in_start(
                            &mut self.chunks,
                            self.current,
                            arr_slot,
                            idx_slot,
                            line,
                        );
                        let elem_slot = self.define_local("__find_elem");
                        self.emit_u16(Op::LOCAL_SET, elem_slot);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u16(Op::LOCAL_GET, elem_slot);
                        self.emit_u8(Op::CALL_REF, 1);
                        {
                            let line = self.line;
                            crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                        };
                        let line = self.line;
                        self.chunk().emit_if(line);
                        self.emit_u16(Op::LOCAL_GET, elem_slot);
                        self.emit_u16(Op::LOCAL_SET, result_slot);
                        self.chunk().emit_br(3, line);
                        self.chunk().emit_end(line);
                        common::loops::emit_for_in_end(
                            &mut self.chunks,
                            self.current,
                            idx_slot,
                            lp,
                            line,
                        );
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                    }
                    "findIndex" | "findindex" => {
                        // findIndex: like find but returns the index, not the element
                        self.emit_const(Value::I32(-1));
                        self.emit_u16(Op::LOCAL_SET, result_slot);
                        let lp = common::loops::emit_for_in_start(
                            &mut self.chunks,
                            self.current,
                            arr_slot,
                            idx_slot,
                            line,
                        );
                        let elem_slot = self.define_local("__findi_elem");
                        self.emit_u16(Op::LOCAL_SET, elem_slot);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u16(Op::LOCAL_GET, elem_slot);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_u8(Op::CALL_REF, 2);
                        {
                            let line = self.line;
                            crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                        };
                        let line = self.line;
                        self.chunk().emit_if(line);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_u16(Op::LOCAL_SET, result_slot);
                        self.chunk().emit_br(3, line);
                        self.chunk().emit_end(line);
                        common::loops::emit_for_in_end(
                            &mut self.chunks,
                            self.current,
                            idx_slot,
                            lp,
                            line,
                        );
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                    }
                    "includes" => {
                        // `x.includes(v[, fromIndex])` — polymorphic:
                        // arrays do element membership, strings do
                        // substring search starting from fromIndex,
                        // user objects fall through to their own
                        // method. Route through `ecma:value.invokeMethod`
                        // so emitted wasm stays spec-compliant.
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        // Pass remaining args (fromIndex etc.) directly
                        // — fn_slot already holds args[0].
                        for extra in arg_exprs.iter().skip(1) {
                            self.compile_expr(extra)?;
                        }
                        common::invoke::emit_invoke_method(
                            &mut self.chunks,
                            self.current,
                            "includes",
                            arg_exprs.len() as u8,
                            line,
                        );
                    }
                    "sort" => {
                        // JS sort(comparatorFn?) — 2-arg comparator or default
                        // ECMA-262 §23.1.3.30: default comparator is
                        // ToString-based ("10" < "2"), not numeric.
                        // Comparator path uses the stdlib (works for JS
                        // and for all other languages); no-comparator JS
                        // routes to ecma:array.sort which does the
                        // spec-compliant lexicographic sort. Other
                        // languages keep stdlib's numeric default.
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit(Op::REF_IS_NULL);
                        let line = self.line;
                        self.chunk().emit_if_value(line);
                        if self.is_js_profile() {
                            // ecma:array.sort returns the sorted array
                            // (in-place, returns receiver). One-arg call.
                            let idx = self.import("ecma:array", "sort");
                            self.emit_u16(Op::LOCAL_GET, arr_slot);
                            self.emit_host_call(idx, 1);
                        } else {
                            let sort_global = self.str_const("__vybe_sort_with_comparator");
                            self.emit_u16(Op::GLOBAL_GET, sort_global);
                            self.emit_u16(Op::LOCAL_GET, arr_slot);
                            self.compile_lambda(
                                &[
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
                                &LambdaBody::Expr(Box::new(Expression::new(ExprKind::Ternary {
                                    cond: Box::new(Expression::new(ExprKind::Binary {
                                        op: BinOp::Lt,
                                        left: Box::new(Expression::ident("left")),
                                        right: Box::new(Expression::ident("right")),
                                    })),
                                    then: Box::new(Expression::new(ExprKind::Lit(Literal::Int(
                                        -1,
                                    )))),
                                    else_: Box::new(Expression::new(ExprKind::Ternary {
                                        cond: Box::new(Expression::new(ExprKind::Binary {
                                            op: BinOp::Gt,
                                            left: Box::new(Expression::ident("left")),
                                            right: Box::new(Expression::ident("right")),
                                        })),
                                        then: Box::new(Expression::new(ExprKind::Lit(
                                            Literal::Int(1),
                                        ))),
                                        else_: Box::new(Expression::new(ExprKind::Lit(
                                            Literal::Int(0),
                                        ))),
                                    })),
                                }))),
                                &[],
                            )?;
                            self.emit_u8(Op::CALL_REF, 2);
                        }
                        self.chunk().emit_else(line);
                        let global = self.str_const("__vybe_sort_with_comparator");
                        self.emit_u16(Op::GLOBAL_GET, global);
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u8(Op::CALL_REF, 2);
                        self.chunk().emit_end(line);
                    }
                    "sort_by_key" => {
                        // .NET OrderBy(keySelector) — 1-arg key extractor
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit(Op::REF_IS_NULL);
                        let line = self.line;
                        self.chunk().emit_if_value(line);
                        let sort_global = self.str_const("__vybe_sort_in_place");
                        self.emit_u16(Op::GLOBAL_GET, sort_global);
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u8(Op::CALL_REF, 1);
                        self.chunk().emit_else(line);
                        let global = self.str_const("__vybe_sort_by_key");
                        self.emit_u16(Op::GLOBAL_GET, global);
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u8(Op::CALL_REF, 2);
                        self.chunk().emit_end(line);
                    }
                    "indexOf" | "indexof" => {
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, fn_slot); // search value
                        common::collections::emit_index_of(&mut self.chunks, self.current, line);
                    }
                    "flatMap" | "flatmap" => {
                        // arr.flatMap(fn) = arr.map(fn).flat()
                        // First emit map: result[i] = fn(arr[i])
                        let mapped_slot = self.define_local("__flatmap_mapped");
                        common::loops::emit_map(
                            &mut self.chunks,
                            self.current,
                            fn_slot,
                            arr_slot,
                            mapped_slot,
                            idx_slot,
                            line,
                        );
                        // Now the mapped array is on stack. Flatten it one level.
                        let flat_idx = self.import("ecma:array", "flat");
                        self.emit_const(Value::I32(1)); // depth = 1
                        self.emit_host_call(flat_idx, 2);
                    }
                    "reduceRight" | "reduceright" => {
                        // reduceRight(fn, initial?) — iterate from end to start.
                        if let Some(init_expr) = arg_exprs.get(1) {
                            self.compile_expr(init_expr)?;
                            self.emit_u16(Op::LOCAL_SET, result_slot);
                        } else {
                            // acc = arr[len-1]
                            self.emit_u16(Op::LOCAL_GET, arr_slot);
                            {
                                let l = self.line;
                                common::collections::emit_len(&mut self.chunks, self.current, l);
                            }
                            self.emit_const(Value::I32(1));
                            self.emit(Op::F64_SUB);
                            self.emit_u16(Op::LOCAL_SET, idx_slot);
                            self.emit_u16(Op::LOCAL_GET, arr_slot);
                            self.emit_u16(Op::LOCAL_GET, idx_slot);
                            {
                                let l = self.line;
                                common::collections::emit_get(&mut self.chunks, self.current, l);
                            }
                            self.emit_u16(Op::LOCAL_SET, result_slot);
                        }
                        // Start from len-1 (or len-2 if no initial)
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        {
                            let l = self.line;
                            common::collections::emit_len(&mut self.chunks, self.current, l);
                        }
                        self.emit_const(Value::I32(1));
                        self.emit(Op::F64_SUB);
                        if arg_exprs.get(1).is_none() {
                            self.emit_const(Value::I32(1));
                            self.emit(Op::F64_SUB);
                        }
                        self.emit_u16(Op::LOCAL_SET, idx_slot);
                        let loop_block = self.chunk().emit_block(line);
                        let (loop_patch, _) = self.chunk().emit_loop_s(line);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_const(Value::I32(0));
                        {
                            let line = self.line;
                            crate::emitter::ops::emit_dyn_ge(self.chunk(), line);
                        };
                        let line = self.line;
                        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                        crate::emitter::ops::emit_dyn_not(self.chunk(), line);
                        self.chunk().emit_br_if(1, line);
                        // acc = fn(acc, arr[i], i)  — ECMA-262 §23.1.3.27
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        {
                            let l = self.line;
                            common::collections::emit_get(&mut self.chunks, self.current, l);
                        }
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_u8(Op::CALL_REF, 3);
                        self.emit_u16(Op::LOCAL_SET, result_slot);
                        // i--
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_const(Value::I32(1));
                        self.emit(Op::F64_SUB);
                        self.emit_u16(Op::LOCAL_SET, idx_slot);
                        self.chunk().emit_br(0, line);
                        self.chunk().emit_end(line);
                        self.chunk().patch_loop(loop_patch);
                        self.chunk().emit_end(line);
                        self.chunk().patch_block(loop_block);
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                    }
                    "findLast" | "findlast" => {
                        // Iterate backward, return last element matching predicate
                        self.emit(Op::NULL);
                        self.emit_u16(Op::LOCAL_SET, result_slot);
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        {
                            let l = self.line;
                            common::collections::emit_len(&mut self.chunks, self.current, l);
                        }
                        self.emit_const(Value::I32(1));
                        self.emit(Op::F64_SUB);
                        self.emit_u16(Op::LOCAL_SET, idx_slot);
                        let loop_block = self.chunk().emit_block(line);
                        let (loop_patch, _) = self.chunk().emit_loop_s(line);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_const(Value::I32(0));
                        {
                            let line = self.line;
                            crate::emitter::ops::emit_dyn_ge(self.chunk(), line);
                        };
                        let line = self.line;
                        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                        crate::emitter::ops::emit_dyn_not(self.chunk(), line);
                        self.chunk().emit_br_if(1, line);
                        let elem_slot = self.define_local("__fl_elem");
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        {
                            let l = self.line;
                            common::collections::emit_get(&mut self.chunks, self.current, l);
                        }
                        self.emit_u16(Op::LOCAL_SET, elem_slot);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u16(Op::LOCAL_GET, elem_slot);
                        self.emit_u8(Op::CALL_REF, 1);
                        {
                            let line = self.line;
                            crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                        };
                        let line = self.line;
                        self.chunk().emit_if(line);
                        self.emit_u16(Op::LOCAL_GET, elem_slot);
                        self.emit_u16(Op::LOCAL_SET, result_slot);
                        self.chunk().emit_br(2, line);
                        self.chunk().emit_end(line);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_const(Value::I32(1));
                        self.emit(Op::F64_SUB);
                        self.emit_u16(Op::LOCAL_SET, idx_slot);
                        self.chunk().emit_br(0, line);
                        self.chunk().emit_end(line);
                        self.chunk().patch_loop(loop_patch);
                        self.chunk().emit_end(line);
                        self.chunk().patch_block(loop_block);
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                    }
                    "findLastIndex" | "findlastindex" => {
                        // Iterate backward, return last index matching predicate
                        self.emit_const(Value::I32(-1));
                        self.emit_u16(Op::LOCAL_SET, result_slot);
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        {
                            let l = self.line;
                            common::collections::emit_len(&mut self.chunks, self.current, l);
                        }
                        self.emit_const(Value::I32(1));
                        self.emit(Op::F64_SUB);
                        self.emit_u16(Op::LOCAL_SET, idx_slot);
                        let loop_block = self.chunk().emit_block(line);
                        let (loop_patch, _) = self.chunk().emit_loop_s(line);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_const(Value::I32(0));
                        {
                            let line = self.line;
                            crate::emitter::ops::emit_dyn_ge(self.chunk(), line);
                        };
                        let line = self.line;
                        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                        crate::emitter::ops::emit_dyn_not(self.chunk(), line);
                        self.chunk().emit_br_if(1, line);
                        let elem_slot2 = self.define_local("__fli_elem");
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        {
                            let l = self.line;
                            common::collections::emit_get(&mut self.chunks, self.current, l);
                        }
                        self.emit_u16(Op::LOCAL_SET, elem_slot2);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u16(Op::LOCAL_GET, elem_slot2);
                        self.emit_u8(Op::CALL_REF, 1);
                        {
                            let line = self.line;
                            crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                        };
                        let line = self.line;
                        self.chunk().emit_if(line);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_u16(Op::LOCAL_SET, result_slot);
                        self.chunk().emit_br(2, line);
                        self.chunk().emit_end(line);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_const(Value::I32(1));
                        self.emit(Op::F64_SUB);
                        self.emit_u16(Op::LOCAL_SET, idx_slot);
                        self.chunk().emit_br(0, line);
                        self.chunk().emit_end(line);
                        self.chunk().patch_loop(loop_patch);
                        self.chunk().emit_end(line);
                        self.chunk().patch_block(loop_block);
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                    }
                    "removeAll" | "removeall" => {
                        // Iterate backward over arr, splice each matching element.
                        // Returns count of removed items.
                        let removed_slot = self.define_local("__ra_removed");
                        self.emit_const(Value::I32(0));
                        self.emit_u16(Op::LOCAL_SET, removed_slot);
                        // Start i = arr.len - 1
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        {
                            let l = self.line;
                            common::collections::emit_len(&mut self.chunks, self.current, l);
                        }
                        self.emit_const(Value::I32(1));
                        self.emit(Op::F64_SUB);
                        self.emit_u16(Op::LOCAL_SET, idx_slot);
                        let loop_block = self.chunk().emit_block(line);
                        let (loop_patch, _) = self.chunk().emit_loop_s(line);
                        // while i >= 0
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_const(Value::I32(0));
                        {
                            let line = self.line;
                            crate::emitter::ops::emit_dyn_ge(self.chunk(), line);
                        };
                        let line = self.line;
                        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                        crate::emitter::ops::emit_dyn_not(self.chunk(), line);
                        self.chunk().emit_br_if(1, line);
                        // elem = arr[i]
                        let ra_elem = self.define_local("__ra_elem");
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        {
                            let l = self.line;
                            common::collections::emit_get(&mut self.chunks, self.current, l);
                        }
                        self.emit_u16(Op::LOCAL_SET, ra_elem);
                        // if fn(elem) → remove
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u16(Op::LOCAL_GET, ra_elem);
                        self.emit_u8(Op::CALL_REF, 1);
                        {
                            let line = self.line;
                            crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                        };
                        let line = self.line;
                        self.chunk().emit_if(line);
                        // splice(arr, i, 1) → drop removed array
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        {
                            let l = self.line;
                            common::collections::emit_remove_at(&mut self.chunks, self.current, l);
                        }
                        self.emit(Op::DROP);
                        // removed++
                        self.emit_u16(Op::LOCAL_GET, removed_slot);
                        self.emit_const(Value::I32(1));
                        {
                            let line = self.line;
                            crate::emitter::ops::emit_dyn_add(self.chunk(), line);
                        };
                        self.emit_u16(Op::LOCAL_SET, removed_slot);
                        self.chunk().emit_end(line);
                        // i--
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_const(Value::I32(1));
                        self.emit(Op::F64_SUB);
                        self.emit_u16(Op::LOCAL_SET, idx_slot);
                        self.chunk().emit_br(0, line);
                        self.chunk().emit_end(line);
                        self.chunk().patch_loop(loop_patch);
                        self.chunk().emit_end(line);
                        self.chunk().patch_block(loop_block);
                        self.emit_u16(Op::LOCAL_GET, removed_slot);
                    }
                    _ => {
                        // Fallback: call as regular method
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u8(Op::CALL_REF, 2);
                    }
                }
                return Ok(());
            }
        }

        // ── Constructor call: ClassName.Create(args) ────────────────
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            if let ExprKind::Ident(class_name) = &object.kind {
                let ctor_nm = &self.profile.constructor_name.clone();
                let is_ctor = if self.case_sensitive {
                    field == ctor_nm
                } else {
                    field.eq_ignore_ascii_case(ctor_nm)
                };
                let canon_class = self.canon(class_name);
                let is_known_class = self.defined_classes.contains(&canon_class)
                    && self.scope().resolve(class_name).is_none();
                if is_ctor && is_known_class {
                    self.emit_var_get(class_name);
                    for a in &arg_exprs {
                        self.compile_expr(a)?;
                    }
                    self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                    return Ok(());
                }
            }
        }

        // ── Pascal builtin helper dispatch: value.Helper(args) ───────
        if self.profile.name == "pascal" {
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                if let Some(type_name) = self.pascal_expr_static_type(object) {
                    let helper_name = self.pascal_helper_function_name(&type_name, field);
                    let helper_canon = self.canon(&helper_name);
                    if self.defined_functions.contains(&helper_canon) {
                        self.emit_var_get(&helper_name);
                        self.compile_expr(object)?;
                        for a in &arg_exprs {
                            self.compile_expr(a)?;
                        }
                        self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                        return Ok(());
                    }

                    let canon_type = self.canon(&type_name);
                    let canon_field = self.canon(field);
                    let is_callable_field = self
                        .pending_classes
                        .get(canon_type.as_str())
                        .map(|pc| pc.fields.iter().any(|name| name == &canon_field))
                        .unwrap_or(false);
                    if is_callable_field {
                        self.compile_expr(object)?;
                        let obj_tmp = self.define_local("__pascal_callable_field_obj");
                        self.emit_u16(Op::LOCAL_SET, obj_tmp);
                        let prop = self.str_const(&canon_field);
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.emit_u16(Op::STRUCT_GET, prop);
                        for a in &arg_exprs {
                            self.compile_expr(a)?;
                        }
                        self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                        return Ok(());
                    }
                }
            }
        }

        // ── Method call: obj.method(args) ───────────────────────────
        if let ExprKind::Member {
            object,
            field,
            null_safe,
        } = &callee.kind
        {
            if self.js_private_member_access_forbidden(field) {
                self.emit_js_private_access_denied(field)?;
                return Ok(());
            }
            if self.is_js_profile() && field.starts_with('#') && !*null_safe {
                self.compile_expr(object)?;
                let obj_tmp = self.define_local("__js_private_call_obj");
                self.emit_u16(Op::LOCAL_SET, obj_tmp);

                let fn_tmp = self.define_local("__js_private_call_fn");
                let class_parts = self.flatten_member_chain(object);
                let static_class_canon = if class_parts.is_empty() {
                    None
                } else {
                    let full_canon = self.canon(&class_parts.join("."));
                    let short_canon =
                        self.canon(class_parts.last().map(String::as_str).unwrap_or(""));
                    if let Some(current_class) = self.current_class.clone() {
                        if class_parts
                            .first()
                            .is_some_and(|part| self.canon(part) == self.canon(&current_class))
                            || class_parts
                                .last()
                                .is_some_and(|part| self.canon(part) == self.canon(&current_class))
                        {
                            Some(current_class)
                        } else if self.defined_classes.contains(&full_canon)
                            || self.pending_classes.contains_key(&full_canon)
                        {
                            Some(full_canon)
                        } else if self.defined_classes.contains(&short_canon)
                            || self.pending_classes.contains_key(&short_canon)
                        {
                            Some(short_canon)
                        } else {
                            None
                        }
                    } else if self.defined_classes.contains(&full_canon)
                        || self.pending_classes.contains_key(&full_canon)
                    {
                        Some(full_canon)
                    } else if self.defined_classes.contains(&short_canon)
                        || self.pending_classes.contains_key(&short_canon)
                    {
                        Some(short_canon)
                    } else {
                        None
                    }
                };

                if let Some(class_name) = static_class_canon {
                    if let Some(overload) =
                        self.resolve_static_method_overload_for_type(&class_name, field, &arg_exprs)
                    {
                        let line = self.line;
                        self.emit_u16(Op::REF_FUNC, overload.chunk_idx as u16);
                        self.chunk().emit(0, line);
                        self.emit_u16(Op::LOCAL_SET, fn_tmp);
                    } else {
                        let field_name = self.js_member_storage_name_for_class(&class_name, field);
                        let prop = self.str_const(&field_name);
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.emit_u16(Op::STRUCT_GET, prop);
                        self.emit_u16(Op::LOCAL_SET, fn_tmp);
                    }
                } else {
                    let field_name = self.js_member_storage_name_for_receiver(object, field);
                    let prop = self.str_const(&field_name);
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit_u16(Op::STRUCT_GET, prop);
                    self.emit_u16(Op::LOCAL_SET, fn_tmp);
                }

                let saved_js_this = self.save_js_this("__js_prev_this_private_call");
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.set_js_this_from_stack();
                self.emit_u16(Op::LOCAL_GET, fn_tmp);
                for arg in &arg_exprs {
                    self.compile_expr(arg)?;
                }
                self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                let result_slot = self.define_local("__js_private_call_result");
                self.emit_u16(Op::LOCAL_SET, result_slot);
                self.restore_js_this(saved_js_this);
                self.emit_u16(Op::LOCAL_GET, result_slot);
                return Ok(());
            }
            if self.is_js_profile() {
                if !*null_safe
                    && self.try_compile_js_promise_chain_call(object, field, &arg_exprs)?
                {
                    return Ok(());
                }

                self.compile_expr(object)?;
                let obj_tmp = self.define_local("__js_obj");
                self.emit_u16(Op::LOCAL_SET, obj_tmp);

                let method_name = self.js_member_storage_name(field);
                let js_result_slot = self.define_local("__js_member_dispatch_result");
                self.emit(Op::NULL);
                self.emit_u16(Op::LOCAL_SET, js_result_slot);
                let js_handled_slot = self.define_local("__js_member_dispatch_handled");
                self.emit_const(Value::I32(0));
                self.emit_u16(Op::LOCAL_SET, js_handled_slot);

                // JS generator objects are both iterators and iterables:
                // `g[Symbol.iterator]()` must return `g` itself.
                if !*null_safe && method_name == "iterator" && arg_exprs.is_empty() {
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    let is_gen_idx = self.import("ecma:value", "isGenerator");
                    self.emit_host_call(is_gen_idx, 1);
                    let gen_if_line = self.line;
                    crate::emitter::ops::emit_dyn_to_bool(self.chunk(), gen_if_line);
                    self.chunk().emit_if(gen_if_line);
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit_u16(Op::LOCAL_SET, js_result_slot);
                    self.emit_const(Value::I32(1));
                    self.emit_u16(Op::LOCAL_SET, js_handled_slot);
                    self.chunk().emit_end(gen_if_line);
                }

                // Generator `.return(v)`: drive the shared generator
                // return-control packet through RESUME so suspended
                // `finally` blocks execute before the completion record
                // is materialized.
                if !*null_safe && method_name == "return" && arg_exprs.len() <= 1 {
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    let is_gen_idx = self.import("ecma:value", "isGenerator");
                    self.emit_host_call(is_gen_idx, 1);
                    let gen_if_line = self.line;
                    crate::emitter::ops::emit_dyn_to_bool(self.chunk(), gen_if_line);
                    self.chunk().emit_if(gen_if_line);

                    let value_slot = self.define_local("__gen_return_value");
                    let done_slot = self.define_local("__gen_return_done");
                    let returned_key = self.str_const("__vybe_gen_returned");

                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    let is_done_idx = self.import("ecma:value", "isGeneratorDone");
                    self.emit_host_call(is_done_idx, 1);
                    let line = self.line;
                    crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                    self.chunk().emit_if(line);

                    if arg_exprs.is_empty() {
                        inst!(self, core_wasm::undefined);
                    } else {
                        self.compile_expr(&arg_exprs[0])?;
                    }
                    self.emit_u16(Op::LOCAL_SET, value_slot);
                    inst!(self, core_wasm::bool_const, true);
                    self.emit_u16(Op::LOCAL_SET, done_slot);
                    self.chunk().emit_else(line);
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    if arg_exprs.is_empty() {
                        inst!(self, core_wasm::undefined);
                    } else {
                        self.compile_expr(&arg_exprs[0])?;
                    }
                    self.emit_generator_control_packet_from_stack("return");
                    let line = self.line;
                    crate::emitter::generators::emit_resume(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_SET, value_slot);

                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    let is_done_idx = self.import("ecma:value", "isGeneratorDone");
                    self.emit_host_call(is_done_idx, 1);
                    self.emit_u16(Op::LOCAL_SET, done_slot);

                    self.chunk().emit_end(line);
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    inst!(self, core_wasm::bool_const, true);
                    self.emit_u16(Op::STRUCT_SET, returned_key);
                    self.emit(Op::DROP);

                    common::dict::emit_new(&mut self.chunks, self.current, line);
                    inst!(self, core_wasm::dup);
                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    let value_key = self.str_const("value");
                    self.emit_u16(Op::STRUCT_SET, value_key);
                    self.emit(Op::DROP);
                    inst!(self, core_wasm::dup);
                    self.emit_u16(Op::LOCAL_GET, done_slot);
                    let done_key = self.str_const("done");
                    self.emit_u16(Op::STRUCT_SET, done_key);
                    self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_SET, js_result_slot);
                    self.emit_const(Value::I32(1));
                    self.emit_u16(Op::LOCAL_SET, js_handled_slot);
                    self.chunk().emit_end(gen_if_line);
                }

                // Generator `.next()` / `.next(v)`: if receiver is a
                // Continuation, drive via WASM stack-switching opcodes
                // and wrap into spec `{value, done}`.
                //   - `g.next()`     → spec `resume` + `(on yield)` handler
                //                       (emit_next; pushes value+has_more)
                //   - `g.next(v)`    → Op::RESUME with v as resume_val
                //                       (pushes yielded value), then
                //                       check `isGeneratorDone` for the
                //                       done flag.
                // Non-Continuations (Array iterators, custom iterables)
                // fall through to regular method dispatch below.
                if !*null_safe && method_name == "next" && arg_exprs.len() <= 1 {
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    let is_gen_idx = self.import("ecma:value", "isGenerator");
                    self.emit_host_call(is_gen_idx, 1);
                    let gen_if_line = self.line;
                    crate::emitter::ops::emit_dyn_to_bool(self.chunk(), gen_if_line);
                    self.chunk().emit_if(gen_if_line);
                    let value_slot = self.define_local("__gen_value");
                    let done_slot = self.define_local("__gen_done");
                    let started_key = self.str_const("__vybe_gen_started");
                    // If a previous `.return()` stamped the cont as
                    // returned, short-circuit to `{value: undefined,
                    // done: true}` per ECMA-262 §27.5.1.2 step 2.
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    let returned_key2 = self.str_const("__vybe_gen_returned");
                    self.emit_u16(Op::STRUCT_GET, returned_key2);
                    {
                        let line = self.line;
                        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    let line = self.line;
                    self.chunk().emit_if(line);
                    inst!(self, core_wasm::undefined);
                    self.emit_u16(Op::LOCAL_SET, value_slot);
                    inst!(self, core_wasm::bool_const, true);
                    self.emit_u16(Op::LOCAL_SET, done_slot);
                    self.chunk().emit_else(line);
                    if arg_exprs.is_empty() {
                        // `g.next()` — GEN_NEXT path: pushes value+has_more.
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        let line = self.line;
                        crate::emitter::generators::emit_next(self.chunk(), line);
                        let has_more_slot = self.define_local("__gen_has_more");
                        self.emit_u16(Op::LOCAL_SET, has_more_slot);
                        self.emit_u16(Op::LOCAL_SET, value_slot);
                        self.emit_u16(Op::LOCAL_GET, has_more_slot);
                        {
                            let line = self.line;
                            crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                        };
                        {
                            let line = self.line;
                            // emit_dyn_not: has_more → i32 (1 if done, 0 if not done)
                            // emit_i32_to_bool: convert to Bool for ECMA `done` property
                            crate::emitter::ops::emit_dyn_not(self.chunk(), line);
                            crate::emitter::ops::emit_i32_to_bool(self.chunk(), line);
                        };
                        self.emit_u16(Op::LOCAL_SET, done_slot);
                        // Per ECMA-262 §27.5.3.5: when a generator completes
                        // (done=true) with no explicit return value, the VM
                        // leaves null on the stack. Convert null → undefined
                        // so the {value} field is spec-correct.
                        if self.is_js_profile() {
                            self.emit_u16(Op::LOCAL_GET, done_slot);
                            let line = self.line;
                            crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                            self.chunk().emit_if(line);
                            self.emit_u16(Op::LOCAL_GET, value_slot);
                            self.emit(Op::REF_IS_NULL);
                            let line = self.line;
                            self.chunk().emit_if(line);
                            inst!(self, core_wasm::undefined);
                            self.emit_u16(Op::LOCAL_SET, value_slot);
                            self.chunk().emit_end(line);
                            self.chunk().emit_end(line);
                        }
                    } else {
                        // `g.next(v)` — RESUME with the resume value;
                        // the suspended yield expression evaluates to
                        // `v`. Pushes only the yielded value back; we
                        // query `isGeneratorDone` for the spec `done`.
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.compile_expr(&arg_exprs[0])?;
                        let line = self.line;
                        crate::emitter::generators::emit_resume(self.chunk(), line);
                        self.emit_u16(Op::LOCAL_SET, value_slot);
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        let is_done_idx = self.import("ecma:value", "isGeneratorDone");
                        self.emit_host_call(is_done_idx, 1);
                        self.emit_u16(Op::LOCAL_SET, done_slot);
                    }
                    self.chunk().emit_end(line);
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    inst!(self, core_wasm::bool_const, true);
                    self.emit_u16(Op::STRUCT_SET, started_key);
                    self.emit(Op::DROP);
                    // Both the early-`returned` short-circuit and the
                    // GEN_NEXT/RESUME paths converge here to build the
                    // `{value, done}` wrapper.
                    common::dict::emit_new(&mut self.chunks, self.current, line);
                    inst!(self, core_wasm::dup);
                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    let value_key = self.str_const("value");
                    self.emit_u16(Op::STRUCT_SET, value_key);
                    self.emit(Op::DROP);
                    inst!(self, core_wasm::dup);
                    self.emit_u16(Op::LOCAL_GET, done_slot);
                    let done_key = self.str_const("done");
                    self.emit_u16(Op::STRUCT_SET, done_key);
                    self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_SET, js_result_slot);
                    self.emit_const(Value::I32(1));
                    self.emit_u16(Op::LOCAL_SET, js_handled_slot);
                    self.chunk().emit_end(gen_if_line);
                }

                if !*null_safe && method_name == "throw" && arg_exprs.len() <= 1 {
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    let is_gen_idx = self.import("ecma:value", "isGenerator");
                    self.emit_host_call(is_gen_idx, 1);
                    let gen_if_line = self.line;
                    crate::emitter::ops::emit_dyn_to_bool(self.chunk(), gen_if_line);
                    self.chunk().emit_if(gen_if_line);

                    let value_slot = self.define_local("__gen_throw_value");
                    let done_slot = self.define_local("__gen_throw_done");
                    let started_key = self.str_const("__vybe_gen_started");

                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit_u16(Op::STRUCT_GET, started_key);
                    {
                        let line = self.line;
                        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.emit(Op::I32_EQZ);
                    let line = self.line;
                    self.chunk().emit_if(line);

                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    let line = self.line;
                    crate::emitter::generators::emit_next(self.chunk(), line);
                    let has_more_slot = self.define_local("__gen_throw_has_more");
                    self.emit_u16(Op::LOCAL_SET, has_more_slot);
                    let primed_value_slot = self.define_local("__gen_throw_primed_value");
                    self.emit_u16(Op::LOCAL_SET, primed_value_slot);

                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    inst!(self, core_wasm::bool_const, true);
                    self.emit_u16(Op::STRUCT_SET, started_key);
                    self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, has_more_slot);
                    {
                        let line = self.line;
                        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.emit(Op::I32_EQZ);
                    let line = self.line;
                    self.chunk().emit_if(line);
                    if arg_exprs.is_empty() {
                        inst!(self, core_wasm::undefined);
                    } else {
                        self.compile_expr(&arg_exprs[0])?;
                    }
                    self.emit(Op::THROW);
                    self.chunk().emit_end(line);
                    self.chunk().emit_end(line);

                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    if arg_exprs.is_empty() {
                        inst!(self, core_wasm::undefined);
                    } else {
                        self.compile_expr(&arg_exprs[0])?;
                    }
                    let line = self.line;
                    crate::emitter::generators::emit_resume_throw(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_SET, value_slot);

                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    let is_done_idx = self.import("ecma:value", "isGeneratorDone");
                    self.emit_host_call(is_done_idx, 1);
                    self.emit_u16(Op::LOCAL_SET, done_slot);

                    common::dict::emit_new(&mut self.chunks, self.current, line);
                    inst!(self, core_wasm::dup);
                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    let value_key = self.str_const("value");
                    self.emit_u16(Op::STRUCT_SET, value_key);
                    self.emit(Op::DROP);
                    inst!(self, core_wasm::dup);
                    self.emit_u16(Op::LOCAL_GET, done_slot);
                    let done_key = self.str_const("done");
                    self.emit_u16(Op::STRUCT_SET, done_key);
                    self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_SET, js_result_slot);
                    self.emit_const(Value::I32(1));
                    self.emit_u16(Op::LOCAL_SET, js_handled_slot);
                    self.chunk().emit_end(gen_if_line);
                }

                let prop = self.str_const(&method_name);
                let receiver_marker = self.str_const("__vybe_method_receiver");
                let js_prefers_typed_dispatch = self
                    .infer_expr_type_hint(object)
                    .as_deref()
                    .map(Self::normalize_type_hint)
                    .is_some_and(|type_hint| js_prefers_typed_member_dispatch(&type_hint));

                if *null_safe {
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit(Op::REF_IS_NULL);
                    let line = self.line;
                    self.chunk().emit_if_value(line);
                    // Per ECMA-262 §13.5.9: optional chain short-circuit yields undefined.
                    inst!(self, core_wasm::undefined);
                    self.chunk().emit_else(line);

                    let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                    for (index, arg) in arg_exprs.iter().enumerate() {
                        self.compile_expr(arg)?;
                        let arg_slot = self.define_local(&format!("__js_lookup_arg_{}", index));
                        self.emit_u16(Op::LOCAL_SET, arg_slot);
                        arg_slots.push(arg_slot);
                    }

                    if js_prefers_typed_dispatch {
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.emit_u16(Op::STRUCT_GET, prop);
                        let fn_slot = self.define_local("__js_typed_method_fn");
                        self.emit_u16(Op::LOCAL_SET, fn_slot);

                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u16(Op::STRUCT_GET, receiver_marker);
                        let marker_slot = self.define_local("__js_typed_receiver_marker");
                        self.emit_u16(Op::LOCAL_SET, marker_slot);

                        self.emit_u16(Op::LOCAL_GET, marker_slot);
                        self.emit(Op::REF_IS_NULL);
                        let line = self.line;
                        self.chunk().emit_if_value(line);
                        self.emit_js_lookup_or_invoke_method_call(
                            obj_tmp,
                            &method_name,
                            &arg_slots,
                        );
                        self.chunk().emit_else(line);
                        self.emit_u16(Op::LOCAL_GET, marker_slot);
                        fn_call!(self, "wasm:js-undefined", "test", 1);
                        let line = self.line;
                        self.chunk().emit_if_value(line);
                        self.emit_js_lookup_or_invoke_method_call(
                            obj_tmp,
                            &method_name,
                            &arg_slots,
                        );
                        self.chunk().emit_else(line);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        for slot in &arg_slots {
                            self.emit_u16(Op::LOCAL_GET, *slot);
                        }
                        self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                        self.chunk().emit_end(line);
                        self.chunk().emit_end(line);
                    } else {
                        self.emit_js_lookup_or_invoke_method_call(
                            obj_tmp,
                            &method_name,
                            &arg_slots,
                        );
                    }
                    self.chunk().emit_end(line);
                    return Ok(());
                }

                if js_prefers_typed_dispatch {
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit_u16(Op::STRUCT_GET, prop);
                    let fn_slot = self.define_local("__js_typed_method_fn");
                    self.emit_u16(Op::LOCAL_SET, fn_slot);

                    self.emit_u16(Op::LOCAL_GET, fn_slot);
                    self.emit_u16(Op::STRUCT_GET, receiver_marker);
                    let marker_slot = self.define_local("__js_typed_receiver_marker");
                    self.emit_u16(Op::LOCAL_SET, marker_slot);

                    self.emit_u16(Op::LOCAL_GET, marker_slot);
                    self.emit(Op::REF_IS_NULL);
                    let line = self.line;
                    self.chunk().emit_if(line);
                    self.chunk().emit_else(line);

                    self.emit_u16(Op::LOCAL_GET, marker_slot);
                    fn_call!(self, "wasm:js-undefined", "test", 1);
                    let line = self.line;
                    self.chunk().emit_if(line);
                    self.chunk().emit_else(line);

                    if args.iter().any(|arg| arg.spread) {
                        let (args_slot, known_len) =
                            self.compile_call_args_array(args, "js_typed_method_spread")?;
                        self.emit_call_ref_with_args_array(
                            fn_slot,
                            Some(obj_tmp),
                            args_slot,
                            known_len,
                        );
                    } else {
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        for a in &arg_exprs {
                            self.compile_expr(a)?;
                        }
                        self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                    }
                    self.emit_u16(Op::LOCAL_SET, js_result_slot);
                    self.emit_const(Value::I32(1));
                    self.emit_u16(Op::LOCAL_SET, js_handled_slot);
                    self.chunk().emit_end(line);
                    self.chunk().emit_end(line);
                }

                self.emit_u16(Op::LOCAL_GET, js_handled_slot);
                self.emit(Op::I32_EQZ);
                let line = self.line;
                self.chunk().emit_if(line);

                let lookup = self.import("ecma:value", "getMethodForCall");
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_const(Value::String(Arc::from(method_name.as_str())));
                self.emit_host_call(lookup, 2);
                let lookup_slot = self.define_local("__js_lookup_fn");
                self.emit_u16(Op::LOCAL_SET, lookup_slot);
                let spread_args = args.iter().any(|arg| arg.spread);
                let spread_call_args = if spread_args {
                    Some(self.compile_call_args_array(args, "js_dispatch_lookup_spread")?)
                } else {
                    None
                };
                let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                if !spread_args {
                    for (index, arg) in arg_exprs.iter().enumerate() {
                        self.compile_expr(arg)?;
                        let arg_slot = self.define_local(&format!("__js_lookup_arg_{}", index));
                        self.emit_u16(Op::LOCAL_SET, arg_slot);
                        arg_slots.push(arg_slot);
                    }
                }
                self.emit_u16(Op::LOCAL_GET, lookup_slot);
                self.emit(Op::REF_IS_NULL);
                let line = self.line;
                self.chunk().emit_if(line);
                if let Some((args_slot, known_len)) = spread_call_args {
                    self.emit_js_invoke_method_from_args_array(
                        obj_tmp,
                        method_name.as_str(),
                        args_slot,
                        known_len,
                    );
                } else {
                    let invoke = self.import("ecma:value", "invokeMethod");
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit_const(Value::String(Arc::from(method_name.as_str())));
                    for slot in &arg_slots {
                        self.emit_u16(Op::LOCAL_GET, *slot);
                    }
                    self.emit_host_call(invoke, (arg_exprs.len() + 2) as u8);
                }
                self.emit_u16(Op::LOCAL_SET, js_result_slot);
                self.chunk().emit_else(line);
                self.emit_u16(Op::LOCAL_GET, lookup_slot);
                fn_call!(self, "wasm:js-undefined", "test", 1);
                let line = self.line;
                self.chunk().emit_if(line);
                if let Some((args_slot, known_len)) = spread_call_args {
                    self.emit_js_invoke_method_from_args_array(
                        obj_tmp,
                        method_name.as_str(),
                        args_slot,
                        known_len,
                    );
                } else {
                    let invoke = self.import("ecma:value", "invokeMethod");
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit_const(Value::String(Arc::from(method_name.as_str())));
                    for slot in &arg_slots {
                        self.emit_u16(Op::LOCAL_GET, *slot);
                    }
                    self.emit_host_call(invoke, (arg_exprs.len() + 2) as u8);
                }
                self.emit_u16(Op::LOCAL_SET, js_result_slot);
                self.chunk().emit_else(line);
                if let Some((args_slot, known_len)) = spread_call_args {
                    self.emit_call_ref_with_args_array(
                        lookup_slot,
                        Some(obj_tmp),
                        args_slot,
                        known_len,
                    );
                } else {
                    self.emit_call_ref_with_bound_js_this_arg_slots(
                        lookup_slot,
                        obj_tmp,
                        &arg_slots,
                    );
                }
                self.emit_u16(Op::LOCAL_SET, js_result_slot);
                self.chunk().emit_end(line);
                self.chunk().emit_end(line);
                self.chunk().emit_end(line);
                self.emit_u16(Op::LOCAL_GET, js_result_slot);
                return Ok(());
            }

            self.compile_expr(object)?;
            let obj_tmp = self.define_local("__obj");
            self.reserve_local_slot(obj_tmp);
            self.emit_u16(Op::LOCAL_SET, obj_tmp);

            if self.is_js_profile() && field.starts_with('#') && !*null_safe {
                let class_parts = self.flatten_member_chain(object);
                let class_name = if let Some(current_class) = self.current_class.clone() {
                    if class_parts
                        .first()
                        .is_some_and(|part| self.canon(part) == self.canon(&current_class))
                        || class_parts
                            .last()
                            .is_some_and(|part| self.canon(part) == self.canon(&current_class))
                    {
                        Some(current_class)
                    } else if !class_parts.is_empty() {
                        let full_canon = self.canon(&class_parts.join("."));
                        let short_canon =
                            self.canon(class_parts.last().map(String::as_str).unwrap_or(""));
                        if self.pending_classes.contains_key(&full_canon)
                            || self.defined_classes.contains(&full_canon)
                        {
                            Some(full_canon)
                        } else if self.pending_classes.contains_key(&short_canon)
                            || self.defined_classes.contains(&short_canon)
                        {
                            Some(short_canon)
                        } else {
                            None
                        }
                    } else {
                        None
                    }
                } else if !class_parts.is_empty() {
                    let full_canon = self.canon(&class_parts.join("."));
                    let short_canon =
                        self.canon(class_parts.last().map(String::as_str).unwrap_or(""));
                    if self.pending_classes.contains_key(&full_canon)
                        || self.defined_classes.contains(&full_canon)
                    {
                        Some(full_canon)
                    } else if self.pending_classes.contains_key(&short_canon)
                        || self.defined_classes.contains(&short_canon)
                    {
                        Some(short_canon)
                    } else {
                        None
                    }
                } else {
                    None
                };

                if let Some(class_name) = class_name {
                    if let Some(chunk_idx) =
                        self.resolve_unique_static_method_chunk_for_class(&class_name, field)
                    {
                        let saved_js_this = self.save_js_this("__js_prev_this_private_static_call");
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.set_js_this_from_stack();
                        let line = self.line;
                        self.emit_u16(Op::REF_FUNC, chunk_idx as u16);
                        self.chunk().emit(0, line);
                        for arg in &arg_exprs {
                            self.compile_expr(arg)?;
                        }
                        self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                        let result_slot = self.define_local("__js_private_static_call_result");
                        self.emit_u16(Op::LOCAL_SET, result_slot);
                        self.restore_js_this(saved_js_this);
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                        return Ok(());
                    }
                }
            }

            let field_name = self.js_member_storage_name_for_receiver(object, field);
            let prop = self.str_const(&field_name);

            if self.profile.parens_for_index && !arg_exprs.is_empty() {
                let is_indexable_typed = self
                    .infer_expr_type_hint(callee)
                    .as_deref()
                    .map(Self::normalize_type_hint)
                    .is_some_and(|type_hint| {
                        Self::is_collection_like_type_hint(&type_hint)
                            && !Self::is_callable_type_hint(&type_hint)
                    });
                if is_indexable_typed {
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit_u16(Op::STRUCT_GET, prop);
                    for arg in &arg_exprs {
                        self.compile_array_index_operand_for_owner(callee, arg)?;
                        let line = self.line;
                        common::collections::emit_get(&mut self.chunks, self.current, line);
                    }
                    return Ok(());
                }
            }

            if *null_safe {
                // obj?.method() — short-circuit to null if obj is null/undefined.
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit(Op::REF_IS_NULL);
                let line = self.line;
                self.chunk().emit_if_value(line);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.chunk().emit_else(line);
                if field.eq_ignore_ascii_case("Invoke") {
                    // C# delegate null-conditional invocation: `d?.Invoke(args)`
                    // should call the delegate value directly when non-null.
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    for a in &arg_exprs {
                        self.compile_expr(a)?;
                    }
                    self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                    self.chunk().emit_end(line);
                    return Ok(());
                }
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, prop);
                let fn_tmp = self.define_local("__fn");
                self.emit_u16(Op::LOCAL_SET, fn_tmp);
                let receiver_key = self.str_const("__vybe_method_receiver");
                self.emit_u16(Op::LOCAL_GET, fn_tmp);
                self.emit_u16(Op::STRUCT_GET, receiver_key);
                let receiver_slot = self.define_local("__member_call_receiver");
                self.emit_u16(Op::LOCAL_SET, receiver_slot);
                let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                for (index, arg) in arg_exprs.iter().enumerate() {
                    self.compile_expr(arg)?;
                    let arg_slot = self.define_local(&format!("__member_nullsafe_arg_{}", index));
                    self.emit_u16(Op::LOCAL_SET, arg_slot);
                    arg_slots.push(arg_slot);
                }
                self.emit_call_ref_with_arg_slots(fn_tmp, Some(receiver_slot), &arg_slots);
                self.chunk().emit_end(line);
                return Ok(());
            }

            let receiver_key = self.str_const("__vybe_method_receiver");

            let buffered_generator_end = if self.profile.buffered_iterator_methods {
                self.emit_buffered_generator_method_dispatch(obj_tmp, &field_name, &arg_exprs)?
            } else {
                None
            };

            if self.is_python_profile() {
                let is_python_generator_method = (field_name == "send" && arg_exprs.len() == 1)
                    || (field_name == "throw" && arg_exprs.len() == 1)
                    || (field_name == "close" && arg_exprs.is_empty());

                if is_python_generator_method {
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    let is_gen_idx = self.import("ecma:value", "isGenerator");
                    self.emit_host_call(is_gen_idx, 1);
                    let line = self.line;
                    crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                    self.chunk().emit_if(line);

                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    match field_name.as_str() {
                        "send" => {
                            self.compile_expr(&arg_exprs[0])?;
                        }
                        "throw" => {
                            self.compile_expr(&arg_exprs[0])?;
                            let line = self.line;
                            crate::emitter::generators::emit_resume_throw(self.chunk(), line);
                        }
                        "close" => {
                            self.emit(Op::NULL);
                            self.emit_generator_control_packet_from_stack("return");
                            let line = self.line;
                            crate::emitter::generators::emit_resume(self.chunk(), line);
                        }
                        _ => unreachable!(),
                    }
                    if field_name == "send" {
                        let line = self.line;
                        crate::emitter::generators::emit_resume(self.chunk(), line);
                    }
                    self.chunk().emit_end(line);
                }
            }

            if let Some(result_slot) = buffered_generator_end {
                if self.profile.namespaces.use_dotnet
                    && arg_exprs.is_empty()
                    && field.eq_ignore_ascii_case("sort")
                    && common::dotnet::uses_runtime_collection_dispatch_arity(field, 0)
                {
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    let line = self.line;
                    self.emit_common("dotnet.array_sort", 1, line);
                    self.finish_buffered_generator_method_dispatch(result_slot);
                    return Ok(());
                }

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, prop);
                let fn_tmp = self.define_local("__fn");
                self.emit_u16(Op::LOCAL_SET, fn_tmp);
                self.emit_u16(Op::LOCAL_GET, fn_tmp);
                self.emit_u16(Op::STRUCT_GET, receiver_key);
                let receiver_slot = self.define_local("__member_fast_receiver");
                self.emit_u16(Op::LOCAL_SET, receiver_slot);
                if self.is_js_profile() {
                    let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                    for (index, arg) in arg_exprs.iter().enumerate() {
                        self.compile_expr(arg)?;
                        let arg_slot =
                            self.define_local(&format!("__js_member_bound_arg_{}", index));
                        self.emit_u16(Op::LOCAL_SET, arg_slot);
                        arg_slots.push(arg_slot);
                    }
                    self.emit_call_ref_with_arg_slots(fn_tmp, None, &arg_slots);
                    self.finish_buffered_generator_method_dispatch(result_slot);
                    return Ok(());
                }
                if self.profile.name == "php" {
                    if let Some(overload) =
                        self.resolve_instance_method_overload(object, field, &arg_exprs, false)
                    {
                        let line = self.line;
                        self.emit_u16(Op::REF_FUNC, overload.chunk_idx as u16);
                        self.chunk().emit(0, line);
                        let direct_fn_tmp = self.define_local("__php_direct_instance_fn");
                        self.emit_u16(Op::LOCAL_SET, direct_fn_tmp);
                        if overload.signature.has_rest {
                            self.emit_known_rest_call_from_local(
                                direct_fn_tmp,
                                Some(obj_tmp),
                                args,
                                &overload.signature,
                            )?;
                        } else {
                            self.emit_u16(Op::LOCAL_GET, direct_fn_tmp);
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            for a in &arg_exprs {
                                self.compile_expr(a)?;
                            }
                            self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                        }
                        self.finish_buffered_generator_method_dispatch(result_slot);
                        return Ok(());
                    }
                }
                if resolves_to_static_container_method(self, object, field) {
                    if self.profile.name == "php" {
                        let class_canon = self.canon(&self.flatten_member_chain(object).join("."));
                        if let Some(overload) = self.resolve_static_method_overload_for_type(
                            &class_canon,
                            field,
                            &arg_exprs,
                        ) {
                            let line = self.line;
                            self.emit_u16(Op::REF_FUNC, overload.chunk_idx as u16);
                            self.chunk().emit(0, line);
                            let direct_fn_tmp = self.define_local("__php_direct_static_fn");
                            self.emit_u16(Op::LOCAL_SET, direct_fn_tmp);
                            if overload.signature.has_rest {
                                self.emit_known_rest_call_from_local(
                                    direct_fn_tmp,
                                    Some(obj_tmp),
                                    args,
                                    &overload.signature,
                                )?;
                            } else {
                                self.emit_u16(Op::LOCAL_GET, direct_fn_tmp);
                                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                                for a in &arg_exprs {
                                    self.compile_expr(a)?;
                                }
                                self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                            }
                            self.finish_buffered_generator_method_dispatch(result_slot);
                            return Ok(());
                        }
                    }
                    if self.profile.name == "csharp" && args.len() == 1 && !args[0].spread {
                        let class_canon = self.canon(&self.flatten_member_chain(object).join("."));
                        if self
                            .resolve_static_method_overload_for_type(
                                &class_canon,
                                field,
                                &arg_exprs,
                            )
                            .is_some_and(|overload| overload.signature.has_rest)
                        {
                            self.emit_variadic_array_call_from_local(fn_tmp, &args[0].value)?;
                            self.finish_buffered_generator_method_dispatch(result_slot);
                            return Ok(());
                        }
                    }
                    if let Some(overload) = self
                        .flatten_member_chain(object)
                        .last()
                        .and_then(|class_name| {
                            self.resolve_static_method_overload_for_type(
                                class_name, field, &arg_exprs,
                            )
                        })
                        .filter(|overload| overload.signature.has_rest)
                    {
                        self.emit_known_rest_call_from_local(
                            fn_tmp,
                            if self.profile.name == "php" {
                                Some(obj_tmp)
                            } else {
                                None
                            },
                            args,
                            &overload.signature,
                        )?;
                        self.finish_buffered_generator_method_dispatch(result_slot);
                        return Ok(());
                    }
                    let rest_signature = self
                        .flatten_member_chain(object)
                        .last()
                        .and_then(|class_name| {
                            self.resolve_static_method_overload_for_type(
                                class_name, field, &arg_exprs,
                            )
                        })
                        .map(|overload| overload.signature.clone())
                        .filter(|signature| signature.has_rest)
                        .or_else(|| {
                            self.function_signatures
                                .get(&self.canon(field))
                                .and_then(|signatures| self.select_call_signature(signatures, args))
                                .filter(|signature| signature.has_rest)
                                .cloned()
                        });
                    if let Some(signature) = rest_signature.as_ref() {
                        self.emit_known_rest_call_from_local(
                            fn_tmp,
                            if self.profile.name == "php" {
                                Some(obj_tmp)
                            } else {
                                None
                            },
                            args,
                            signature,
                        )?;
                    } else if self.is_js_profile() {
                        let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                        for (index, arg) in arg_exprs.iter().enumerate() {
                            self.compile_expr(arg)?;
                            let arg_slot =
                                self.define_local(&format!("__js_static_member_arg_{}", index));
                            self.emit_u16(Op::LOCAL_SET, arg_slot);
                            arg_slots.push(arg_slot);
                        }
                        self.emit_call_ref_with_arg_slots(fn_tmp, None, &arg_slots);
                    } else {
                        let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                        for (index, arg) in arg_exprs.iter().enumerate() {
                            self.compile_expr(arg)?;
                            let arg_slot =
                                self.define_local(&format!("__static_member_arg_{}", index));
                            self.emit_u16(Op::LOCAL_SET, arg_slot);
                            arg_slots.push(arg_slot);
                        }
                        self.emit_call_ref_with_arg_slots(
                            fn_tmp,
                            if self.profile.name == "php" {
                                Some(obj_tmp)
                            } else {
                                None
                            },
                            &arg_slots,
                        );
                    }
                    if args.iter().any(|arg| arg.by_ref) {
                        let pack_slot = self.define_local("__member_static_container_by_ref_pack");
                        self.emit_u16(Op::LOCAL_SET, pack_slot);
                        let mut ref_out_index = 1usize;
                        for arg in args {
                            if !arg.by_ref {
                                continue;
                            }
                            self.emit_u16(Op::LOCAL_GET, pack_slot);
                            self.emit_const(Value::F64(ref_out_index as f64));
                            common::collections::emit_get(
                                &mut self.chunks,
                                self.current,
                                self.line,
                            );
                            self.compile_assign_target(&arg.value)?;
                            ref_out_index += 1;
                        }
                        self.emit_u16(Op::LOCAL_GET, pack_slot);
                        self.emit_const(Value::F64(0.0));
                        common::collections::emit_get(&mut self.chunks, self.current, self.line);
                    }
                    self.finish_buffered_generator_method_dispatch(result_slot);
                    return Ok(());
                }
                let primitive_tostring_if = if self.profile.namespaces.use_dotnet
                    && arg_exprs.is_empty()
                    && field.eq_ignore_ascii_case("ToString")
                {
                    let type_tmp = self.define_local("__dotnet_tostring_type");
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    fn_call!(self, "ecma:value", "typeof", 1);
                    self.emit_u16(Op::LOCAL_SET, type_tmp);

                    let primitive_slot = self.define_local("__dotnet_tostring_is_primitive");
                    self.emit_const(Value::I32(0));
                    self.emit_u16(Op::LOCAL_SET, primitive_slot);
                    for type_name in ["number", "i32", "i64", "string", "boolean"] {
                        self.emit_u16(Op::LOCAL_GET, type_tmp);
                        self.emit_const(Value::String(Arc::from(type_name)));
                        {
                            let line = self.line;
                            crate::emitter::ops::emit_dyn_eq(self.chunk(), line);
                        };
                        let line = self.line;
                        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                        self.chunk().emit_if(line);
                        self.emit_const(Value::I32(1));
                        self.emit_u16(Op::LOCAL_SET, primitive_slot);
                        self.chunk().emit_end(line);
                    }

                    self.emit_u16(Op::LOCAL_GET, primitive_slot);
                    let line = self.line;
                    self.chunk().emit_if_value(line);
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    let line = self.line;
                    common::strings::emit_to_string(self.chunk(), line);
                    self.chunk().emit_else(line);
                    Some(line)
                } else {
                    None
                };
                if self.profile.namespaces.use_dotnet
                    && arg_exprs.is_empty()
                    && field.eq_ignore_ascii_case("Count")
                {
                    self.emit_u16(Op::LOCAL_GET, fn_tmp);
                    fn_call!(self, "ecma:value", "typeof", 1);
                    self.emit_const(Value::String(Arc::from("function")));
                    {
                        let line = self.line;
                        crate::emitter::ops::emit_dyn_eq(self.chunk(), line);
                    };
                    let line = self.line;
                    crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                    self.chunk().emit_if_value(line);

                    self.emit_u16(Op::LOCAL_GET, fn_tmp);
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit_u8(Op::CALL_REF, 1);
                    self.chunk().emit_else(line);
                    self.emit_u16(Op::LOCAL_GET, fn_tmp);
                    self.chunk().emit_end(line);
                    self.finish_buffered_generator_method_dispatch(result_slot);
                    return Ok(());
                }
                if self.is_js_profile() {
                    let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                    for (index, arg) in arg_exprs.iter().enumerate() {
                        self.compile_expr(arg)?;
                        let arg_slot =
                            self.define_local(&format!("__js_member_fast_arg_{}", index));
                        self.emit_u16(Op::LOCAL_SET, arg_slot);
                        arg_slots.push(arg_slot);
                    }
                    if field == "toString" {
                        let invoke = self.import("ecma:value", "invokeMethod");
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.emit_const(Value::String(Arc::from(field.as_str())));
                        for slot in &arg_slots {
                            self.emit_u16(Op::LOCAL_GET, *slot);
                        }
                        self.emit_host_call(invoke, (arg_slots.len() + 2) as u8);
                    } else {
                        self.emit_call_ref_with_arg_slots(fn_tmp, None, &arg_slots);
                    }
                } else {
                    if let Some(param_modes) =
                        self.function_param_modes.get(&self.canon(field)).cloned()
                    {
                        let receiver_param_offset =
                            usize::from(param_modes.len() == args.len() + 1);
                        let user_modes =
                            &param_modes[receiver_param_offset.min(param_modes.len())..];
                        if user_modes
                            .iter()
                            .any(|mode| matches!(mode, PassBy::Ref | PassBy::Out))
                        {
                            let mut arg_slots = Vec::with_capacity(args.len());
                            for (index, arg) in args.iter().enumerate() {
                                self.compile_ref_aware_call_arg(
                                    arg,
                                    user_modes.get(index).copied().unwrap_or(PassBy::Value),
                                )?;
                                let arg_slot =
                                    self.define_local(&format!("__member_fast_arg_{}", index));
                                self.emit_u16(Op::LOCAL_SET, arg_slot);
                                arg_slots.push(arg_slot);
                            }

                            self.emit_call_ref_with_arg_slots(
                                fn_tmp,
                                Some(receiver_slot),
                                &arg_slots,
                            );

                            let pack_slot = self.define_local("__member_fast_ref_call_pack");
                            self.emit_u16(Op::LOCAL_SET, pack_slot);
                            let mut ref_out_index = 1usize;
                            for (index, arg) in args.iter().enumerate() {
                                if !matches!(user_modes.get(index), Some(PassBy::Ref | PassBy::Out))
                                {
                                    continue;
                                }
                                self.emit_u16(Op::LOCAL_GET, pack_slot);
                                self.emit_const(Value::F64(ref_out_index as f64));
                                common::collections::emit_get(
                                    &mut self.chunks,
                                    self.current,
                                    self.line,
                                );
                                self.compile_assign_target(&arg.value)?;
                                ref_out_index += 1;
                            }
                            self.emit_u16(Op::LOCAL_GET, pack_slot);
                            self.emit_const(Value::F64(0.0));
                            common::collections::emit_get(
                                &mut self.chunks,
                                self.current,
                                self.line,
                            );
                            self.emit_u16(Op::LOCAL_GET, receiver_slot);
                            self.emit(Op::REF_IS_NULL);
                            let line = self.line;
                            self.chunk().emit_if(line);
                            self.chunk().emit_else(line);
                            self.emit_u16(Op::LOCAL_GET, receiver_slot);
                            fn_call!(self, "wasm:js-undefined", "test", 1);
                            let line = self.line;
                            self.chunk().emit_if(line);
                            self.chunk().emit_else(line);
                            self.emit_fortran_member_receiver_writeback(object, obj_tmp);
                            self.chunk().emit_end(line);
                            self.chunk().emit_end(line);
                        } else {
                            let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                            for (index, arg) in arg_exprs.iter().enumerate() {
                                self.compile_expr(arg)?;
                                let arg_slot =
                                    self.define_local(&format!("__member_fast_arg_{}", index));
                                self.emit_u16(Op::LOCAL_SET, arg_slot);
                                arg_slots.push(arg_slot);
                            }
                            self.emit_call_ref_with_arg_slots(
                                fn_tmp,
                                Some(receiver_slot),
                                &arg_slots,
                            );
                            self.emit_u16(Op::LOCAL_GET, receiver_slot);
                            self.emit(Op::REF_IS_NULL);
                            let line = self.line;
                            self.chunk().emit_if(line);
                            self.chunk().emit_else(line);
                            self.emit_u16(Op::LOCAL_GET, receiver_slot);
                            fn_call!(self, "wasm:js-undefined", "test", 1);
                            let line = self.line;
                            self.chunk().emit_if(line);
                            self.chunk().emit_else(line);
                            self.emit_fortran_member_receiver_writeback(object, obj_tmp);
                            self.chunk().emit_end(line);
                            self.chunk().emit_end(line);
                        }
                    } else {
                        let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                        for (index, arg) in arg_exprs.iter().enumerate() {
                            self.compile_expr(arg)?;
                            let arg_slot =
                                self.define_local(&format!("__member_fast_arg_{}", index));
                            self.emit_u16(Op::LOCAL_SET, arg_slot);
                            arg_slots.push(arg_slot);
                        }
                        self.emit_call_ref_with_arg_slots(fn_tmp, Some(receiver_slot), &arg_slots);
                        self.emit_u16(Op::LOCAL_GET, receiver_slot);
                        self.emit(Op::REF_IS_NULL);
                        let line = self.line;
                        self.chunk().emit_if(line);
                        self.chunk().emit_else(line);
                        self.emit_u16(Op::LOCAL_GET, receiver_slot);
                        fn_call!(self, "wasm:js-undefined", "test", 1);
                        let line = self.line;
                        self.chunk().emit_if(line);
                        self.chunk().emit_else(line);
                        self.emit_fortran_member_receiver_writeback(object, obj_tmp);
                        self.chunk().emit_end(line);
                        self.chunk().emit_end(line);
                    }
                }
                if let Some(line) = primitive_tostring_if {
                    self.chunk().emit_end(line);
                }
                self.finish_buffered_generator_method_dispatch(result_slot);
                return Ok(());
            }

            if self.profile.namespaces.use_dotnet
                && arg_exprs.is_empty()
                && field.eq_ignore_ascii_case("sort")
                && common::dotnet::uses_runtime_collection_dispatch_arity(field, 0)
            {
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                let line = self.line;
                self.emit_common("dotnet.array_sort", 1, line);
                return Ok(());
            }

            if let Some(chunk_idx) =
                self.resolve_instance_method_overload_chunk(object, field, &arg_exprs)
            {
                let overload = self
                    .resolve_instance_method_overload(object, field, &arg_exprs, true)
                    .ok_or_else(|| format!("failed to resolve method overload for {}", field))?;
                if overload.signature.has_rest {
                    let line = self.line;
                    self.emit_u16(Op::REF_FUNC, chunk_idx as u16);
                    self.chunk().emit(0, line);
                    let direct_fn_tmp = self.define_local("__direct_instance_fn");
                    self.emit_u16(Op::LOCAL_SET, direct_fn_tmp);
                    self.emit_known_rest_call_from_local(
                        direct_fn_tmp,
                        if self.is_js_profile() {
                            None
                        } else {
                            Some(obj_tmp)
                        },
                        args,
                        &overload.signature,
                    )?;
                } else {
                    self.emit_direct_instance_method_call(
                        chunk_idx, field, obj_tmp, args, &arg_exprs,
                    )?;
                }
                return Ok(());
            }

            if let Some(class_name) = resolve_go_pending_instance_method_owner(self, object, field)
            {
                let class_idx = self.str_const(&class_name);
                self.emit_u16(Op::GLOBAL_GET, class_idx);
                self.emit_u16(Op::STRUCT_GET, prop);
                let fn_tmp = self.define_local("__go_pending_instance_fn");
                self.emit_u16(Op::LOCAL_SET, fn_tmp);

                let receiver_slot =
                    if self
                        .pending_classes
                        .get(&class_name)
                        .is_some_and(|pending| {
                            !pending
                                .instance_pointer_method_names
                                .iter()
                                .any(|name| self.canon(name) == self.canon(field))
                                && !pending.fields.is_empty()
                        })
                    {
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.emit_autoderef_pointer_cell();
                        self.emit_user_value_type_clone_from_stack(&class_name);
                        let receiver_slot = self.define_local("__go_value_receiver");
                        self.emit_u16(Op::LOCAL_SET, receiver_slot);
                        receiver_slot
                    } else {
                        obj_tmp
                    };

                let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                for (index, arg) in arg_exprs.iter().enumerate() {
                    self.compile_expr(arg)?;
                    let arg_slot = self.define_local(&format!("__go_pending_method_arg_{}", index));
                    self.emit_u16(Op::LOCAL_SET, arg_slot);
                    arg_slots.push(arg_slot);
                }
                self.emit_call_ref_with_arg_slots(fn_tmp, Some(receiver_slot), &arg_slots);
                return Ok(());
            }

            self.emit_u16(Op::LOCAL_GET, obj_tmp);
            self.emit_autoderef_pointer_cell();
            self.emit_u16(Op::STRUCT_GET, prop);
            let fn_tmp = self.define_local("__fn");
            self.emit_u16(Op::LOCAL_SET, fn_tmp);
            if self.profile.name == "php" {
                if let Some(overload) =
                    self.resolve_instance_method_overload(object, field, &arg_exprs, false)
                {
                    let line = self.line;
                    self.emit_u16(Op::REF_FUNC, overload.chunk_idx as u16);
                    self.chunk().emit(0, line);
                    let direct_fn_tmp = self.define_local("__php_direct_instance_fn");
                    self.emit_u16(Op::LOCAL_SET, direct_fn_tmp);
                    if overload.signature.has_rest {
                        self.emit_known_rest_call_from_local(
                            direct_fn_tmp,
                            Some(obj_tmp),
                            args,
                            &overload.signature,
                        )?;
                    } else {
                        self.emit_u16(Op::LOCAL_GET, direct_fn_tmp);
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        for a in &arg_exprs {
                            self.compile_expr(a)?;
                        }
                        self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                    }
                    return Ok(());
                }
            }
            let member_index_if = if self.profile.parens_for_index && !arg_exprs.is_empty() {
                self.emit_u16(Op::LOCAL_GET, fn_tmp);
                fn_call!(self, "ecma:value", "typeof", 1);
                self.emit_const(Value::String(Arc::from("function")));
                {
                    let line = self.line;
                    crate::emitter::ops::emit_dyn_eq(self.chunk(), line);
                };
                let line = self.line;
                crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                self.chunk().emit_if_value(line);
                Some(line)
            } else {
                None
            };
            if resolves_to_static_container_method(self, object, field) {
                if self.profile.name == "php" {
                    let class_canon = self.canon(&self.flatten_member_chain(object).join("."));
                    if let Some(overload) = self.resolve_static_method_overload_for_type(
                        &class_canon,
                        field,
                        &arg_exprs,
                    ) {
                        let line = self.line;
                        self.emit_u16(Op::REF_FUNC, overload.chunk_idx as u16);
                        self.chunk().emit(0, line);
                        let direct_fn_tmp = self.define_local("__php_direct_static_fn");
                        self.emit_u16(Op::LOCAL_SET, direct_fn_tmp);
                        if overload.signature.has_rest {
                            self.emit_known_rest_call_from_local(
                                direct_fn_tmp,
                                Some(obj_tmp),
                                args,
                                &overload.signature,
                            )?;
                        } else {
                            self.emit_u16(Op::LOCAL_GET, direct_fn_tmp);
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            for a in &arg_exprs {
                                self.compile_expr(a)?;
                            }
                            self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                        }
                        if let Some(line) = member_index_if {
                            self.finish_member_index_call_path(callee, &arg_exprs, fn_tmp, line)?;
                        }
                        return Ok(());
                    }
                }
                if self.profile.name == "csharp" && args.len() == 1 && !args[0].spread {
                    let class_canon = self.canon(&self.flatten_member_chain(object).join("."));
                    if self
                        .resolve_static_method_overload_for_type(&class_canon, field, &arg_exprs)
                        .is_some_and(|overload| overload.signature.has_rest)
                    {
                        self.emit_variadic_array_call_from_local(fn_tmp, &args[0].value)?;
                        if let Some(line) = member_index_if {
                            self.finish_member_index_call_path(callee, &arg_exprs, fn_tmp, line)?;
                        }
                        return Ok(());
                    }
                }
                let static_overload =
                    self.flatten_member_chain(object)
                        .last()
                        .and_then(|class_name| {
                            self.resolve_static_method_overload_for_type(
                                class_name, field, &arg_exprs,
                            )
                        });
                let rest_signature = static_overload
                    .as_ref()
                    .map(|overload| overload.signature.clone())
                    .filter(|signature| signature.has_rest)
                    .or_else(|| {
                        self.function_signatures
                            .get(&self.canon(field))
                            .and_then(|signatures| self.select_call_signature(signatures, args))
                            .filter(|signature| signature.has_rest)
                            .cloned()
                    });
                if let Some(signature) = rest_signature.as_ref() {
                    let line = self.line;
                    let rest_callee_slot = if let Some(overload) = static_overload.as_ref() {
                        self.emit_u16(Op::REF_FUNC, overload.chunk_idx as u16);
                        self.chunk().emit(0, line);
                        let direct_fn_tmp = self.define_local("__static_rest_body_fn");
                        self.emit_u16(Op::LOCAL_SET, direct_fn_tmp);
                        direct_fn_tmp
                    } else {
                        fn_tmp
                    };
                    self.emit_known_rest_call_from_local(
                        rest_callee_slot,
                        if self.profile.name == "php" {
                            Some(obj_tmp)
                        } else {
                            None
                        },
                        args,
                        signature,
                    )?;
                } else {
                    let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                    for (index, arg) in arg_exprs.iter().enumerate() {
                        self.compile_expr(arg)?;
                        let arg_slot =
                            self.define_local(&format!("__static_member_call_arg_{}", index));
                        self.emit_u16(Op::LOCAL_SET, arg_slot);
                        arg_slots.push(arg_slot);
                    }
                    self.emit_call_ref_with_arg_slots(
                        fn_tmp,
                        if self.profile.name == "php" {
                            Some(obj_tmp)
                        } else {
                            None
                        },
                        &arg_slots,
                    );
                }
                if args.iter().any(|arg| arg.by_ref) {
                    let pack_slot = self.define_local("__member_static_container_pack");
                    self.emit_u16(Op::LOCAL_SET, pack_slot);
                    let mut ref_out_index = 1usize;
                    for arg in args {
                        if !arg.by_ref {
                            continue;
                        }
                        self.emit_u16(Op::LOCAL_GET, pack_slot);
                        self.emit_const(Value::F64(ref_out_index as f64));
                        common::collections::emit_get(&mut self.chunks, self.current, self.line);
                        self.compile_assign_target(&arg.value)?;
                        ref_out_index += 1;
                    }
                    self.emit_u16(Op::LOCAL_GET, pack_slot);
                    self.emit_const(Value::F64(0.0));
                    common::collections::emit_get(&mut self.chunks, self.current, self.line);
                }
                if let Some(line) = member_index_if {
                    self.finish_member_index_call_path(callee, &arg_exprs, fn_tmp, line)?;
                }
                return Ok(());
            }
            let primitive_tostring_if = if self.profile.namespaces.use_dotnet
                && arg_exprs.is_empty()
                && field.eq_ignore_ascii_case("ToString")
            {
                let type_tmp = self.define_local("__dotnet_tostring_type");
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                fn_call!(self, "ecma:value", "typeof", 1);
                self.emit_u16(Op::LOCAL_SET, type_tmp);

                let primitive_slot = self.define_local("__dotnet_tostring_is_primitive");
                self.emit_const(Value::I32(0));
                self.emit_u16(Op::LOCAL_SET, primitive_slot);
                for type_name in ["number", "i32", "i64", "string", "boolean"] {
                    self.emit_u16(Op::LOCAL_GET, type_tmp);
                    self.emit_const(Value::String(Arc::from(type_name)));
                    {
                        let line = self.line;
                        crate::emitter::ops::emit_dyn_eq(self.chunk(), line);
                    };
                    let line = self.line;
                    crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                    self.chunk().emit_if(line);
                    self.emit_const(Value::I32(1));
                    self.emit_u16(Op::LOCAL_SET, primitive_slot);
                    self.chunk().emit_end(line);
                }

                self.emit_u16(Op::LOCAL_GET, primitive_slot);
                let line = self.line;
                self.chunk().emit_if_value(line);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                let line = self.line;
                common::strings::emit_to_string(self.chunk(), line);
                self.chunk().emit_else(line);
                Some(line)
            } else {
                None
            };
            if self.profile.namespaces.use_dotnet
                && arg_exprs.is_empty()
                && field.eq_ignore_ascii_case("Count")
            {
                self.emit_u16(Op::LOCAL_GET, fn_tmp);
                fn_call!(self, "ecma:value", "typeof", 1);
                self.emit_const(Value::String(Arc::from("function")));
                {
                    let line = self.line;
                    crate::emitter::ops::emit_dyn_eq(self.chunk(), line);
                };
                let line = self.line;
                crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                self.chunk().emit_if_value(line);

                self.emit_u16(Op::LOCAL_GET, fn_tmp);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u8(Op::CALL_REF, 1);
                self.chunk().emit_else(line);
                self.emit_u16(Op::LOCAL_GET, fn_tmp);
                self.chunk().emit_end(line);
                return Ok(());
            }
            if let Some(overload) = self
                .resolve_instance_method_overload(object, field, &arg_exprs, false)
                .filter(|overload| overload.signature.has_rest)
            {
                let line = self.line;
                self.emit_u16(Op::REF_FUNC, overload.chunk_idx as u16);
                self.chunk().emit(0, line);
                let direct_fn_tmp = self.define_local("__instance_rest_body_fn");
                self.emit_u16(Op::LOCAL_SET, direct_fn_tmp);
                self.emit_known_rest_call_from_local(
                    direct_fn_tmp,
                    if self.is_js_profile() {
                        None
                    } else {
                        Some(obj_tmp)
                    },
                    args,
                    &overload.signature,
                )?;
            } else {
                let receiver_slot = obj_tmp;
                if self.is_js_profile() {
                    let js_result_slot = self.define_local("__js_member_result");
                    self.emit(Op::NULL);
                    self.emit_u16(Op::LOCAL_SET, js_result_slot);
                    let js_handled_slot = self.define_local("__js_member_handled");
                    self.emit_const(Value::I32(0));
                    self.emit_u16(Op::LOCAL_SET, js_handled_slot);

                    let js_user_defined_member =
                        self.direct_receiver_has_own_pending_method(object, field)
                            || self.infer_expr_type_hint(object).as_deref().is_some_and(
                                |type_hint| {
                                    self.pending_class_has_method_name_for_type(type_hint, field)
                                },
                            );
                    if js_user_defined_member {
                        self.emit_u16(Op::LOCAL_GET, fn_tmp);
                        self.emit(Op::REF_IS_NULL);
                        let line = self.line;
                        self.chunk().emit_if(line);
                        self.chunk().emit_else(line);
                        self.emit_u16(Op::LOCAL_GET, fn_tmp);
                        fn_call!(self, "wasm:js-undefined", "test", 1);
                        let line = self.line;
                        self.chunk().emit_if(line);
                        self.chunk().emit_else(line);

                        if args.iter().any(|arg| arg.spread) {
                            let (args_slot, known_len) =
                                self.compile_call_args_array(args, "js_member_call_spread")?;
                            self.emit_call_ref_with_args_array(
                                fn_tmp,
                                Some(obj_tmp),
                                args_slot,
                                known_len,
                            );
                        } else {
                            let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                            for (index, arg) in arg_exprs.iter().enumerate() {
                                self.compile_expr(arg)?;
                                let arg_slot =
                                    self.define_local(&format!("__js_member_call_arg_{}", index));
                                self.emit_u16(Op::LOCAL_SET, arg_slot);
                                arg_slots.push(arg_slot);
                            }
                            self.emit_call_ref_with_bound_js_this_arg_slots(
                                fn_tmp, obj_tmp, &arg_slots,
                            );
                        }
                        self.emit_u16(Op::LOCAL_SET, js_result_slot);
                        self.emit_const(Value::I32(1));
                        self.emit_u16(Op::LOCAL_SET, js_handled_slot);
                        self.chunk().emit_end(line);
                        self.chunk().emit_end(line);
                    }

                    self.emit_u16(Op::LOCAL_GET, js_handled_slot);
                    self.emit(Op::I32_EQZ);
                    let line = self.line;
                    self.chunk().emit_if(line);
                    let method_name = field_name.clone();
                    let lookup = self.import("ecma:value", "getMethodForCall");
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit_const(Value::String(Arc::from(method_name.as_str())));
                    self.emit_host_call(lookup, 2);
                    let lookup_slot = self.define_local("__js_member_lookup_fn");
                    self.emit_u16(Op::LOCAL_SET, lookup_slot);
                    let spread_args = args.iter().any(|arg| arg.spread);
                    let spread_call_args = if spread_args {
                        Some(self.compile_call_args_array(args, "js_member_lookup_spread")?)
                    } else {
                        None
                    };
                    let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                    if !spread_args {
                        for (index, arg) in arg_exprs.iter().enumerate() {
                            self.compile_expr(arg)?;
                            let arg_slot =
                                self.define_local(&format!("__js_member_lookup_arg_{}", index));
                            self.emit_u16(Op::LOCAL_SET, arg_slot);
                            arg_slots.push(arg_slot);
                        }
                    }

                    self.emit_u16(Op::LOCAL_GET, lookup_slot);
                    self.emit(Op::REF_IS_NULL);
                    let line = self.line;
                    self.chunk().emit_if(line);
                    if let Some((args_slot, known_len)) = spread_call_args {
                        self.emit_js_invoke_method_from_args_array(
                            obj_tmp,
                            method_name.as_str(),
                            args_slot,
                            known_len,
                        );
                    } else {
                        let invoke = self.import("ecma:value", "invokeMethod");
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.emit_const(Value::String(Arc::from(method_name.as_str())));
                        for slot in &arg_slots {
                            self.emit_u16(Op::LOCAL_GET, *slot);
                        }
                        self.emit_host_call(invoke, (arg_slots.len() + 2) as u8);
                    }
                    self.emit_u16(Op::LOCAL_SET, js_result_slot);
                    self.chunk().emit_else(line);
                    self.emit_u16(Op::LOCAL_GET, lookup_slot);
                    fn_call!(self, "wasm:js-undefined", "test", 1);
                    let line = self.line;
                    self.chunk().emit_if(line);
                    if let Some((args_slot, known_len)) = spread_call_args {
                        self.emit_js_invoke_method_from_args_array(
                            obj_tmp,
                            method_name.as_str(),
                            args_slot,
                            known_len,
                        );
                    } else {
                        let invoke = self.import("ecma:value", "invokeMethod");
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.emit_const(Value::String(Arc::from(method_name.as_str())));
                        for slot in &arg_slots {
                            self.emit_u16(Op::LOCAL_GET, *slot);
                        }
                        self.emit_host_call(invoke, (arg_slots.len() + 2) as u8);
                    }
                    self.emit_u16(Op::LOCAL_SET, js_result_slot);
                    self.chunk().emit_else(line);
                    if let Some((args_slot, known_len)) = spread_call_args {
                        self.emit_call_ref_with_args_array(
                            lookup_slot,
                            Some(obj_tmp),
                            args_slot,
                            known_len,
                        );
                    } else {
                        self.emit_call_ref_with_bound_js_this_arg_slots(
                            lookup_slot,
                            obj_tmp,
                            &arg_slots,
                        );
                    }
                    self.emit_u16(Op::LOCAL_SET, js_result_slot);
                    self.chunk().emit_end(line);
                    self.chunk().emit_end(line);
                    self.chunk().emit_end(line);
                    self.emit_u16(Op::LOCAL_GET, js_result_slot);
                } else {
                    let method_canon = self.canon(field);
                    let qualified_method = resolve_receiver_type_hint(self, object)
                        .and_then(|receiver_type| {
                            self.resolve_pending_class_name_for_type_hint(&receiver_type)
                        })
                        .map(|class_name| self.canon(&format!("{}.{}", class_name, field)));
                    if let Some(param_modes) = qualified_method
                        .as_ref()
                        .and_then(|qualified| self.function_param_modes.get(qualified).cloned())
                        .or_else(|| self.function_param_modes.get(&method_canon).cloned())
                    {
                        let receiver_param_offset =
                            usize::from(param_modes.len() == args.len() + 1);
                        let user_modes =
                            &param_modes[receiver_param_offset.min(param_modes.len())..];
                        if user_modes
                            .iter()
                            .any(|mode| matches!(mode, PassBy::Ref | PassBy::Out))
                        {
                            let mut arg_slots = Vec::with_capacity(args.len());
                            for (index, arg) in args.iter().enumerate() {
                                self.compile_ref_aware_call_arg(
                                    arg,
                                    user_modes.get(index).copied().unwrap_or(PassBy::Value),
                                )?;
                                let arg_slot =
                                    self.define_local(&format!("__member_call_arg_{}", index));
                                self.emit_u16(Op::LOCAL_SET, arg_slot);
                                arg_slots.push(arg_slot);
                            }
                            self.emit_call_ref_with_arg_slots(
                                fn_tmp,
                                Some(receiver_slot),
                                &arg_slots,
                            );

                            let pack_slot = self.define_local("__member_call_ref_pack");
                            self.emit_u16(Op::LOCAL_SET, pack_slot);
                            let mut ref_out_index = 1usize;
                            for (index, arg) in args.iter().enumerate() {
                                if !matches!(user_modes.get(index), Some(PassBy::Ref | PassBy::Out))
                                {
                                    continue;
                                }
                                self.emit_u16(Op::LOCAL_GET, pack_slot);
                                self.emit_const(Value::F64(ref_out_index as f64));
                                common::collections::emit_get(
                                    &mut self.chunks,
                                    self.current,
                                    self.line,
                                );
                                self.compile_assign_target(&arg.value)?;
                                ref_out_index += 1;
                            }
                            self.emit_u16(Op::LOCAL_GET, pack_slot);
                            self.emit_const(Value::F64(0.0));
                            common::collections::emit_get(
                                &mut self.chunks,
                                self.current,
                                self.line,
                            );
                            self.emit_u16(Op::LOCAL_GET, receiver_slot);
                            self.emit(Op::REF_IS_NULL);
                            let line = self.line;
                            self.chunk().emit_if(line);
                            self.chunk().emit_else(line);
                            self.emit_u16(Op::LOCAL_GET, receiver_slot);
                            fn_call!(self, "wasm:js-undefined", "test", 1);
                            let line = self.line;
                            self.chunk().emit_if(line);
                            self.chunk().emit_else(line);
                            self.emit_fortran_member_receiver_writeback(object, obj_tmp);
                            self.chunk().emit_end(line);
                            self.chunk().emit_end(line);
                        } else {
                            let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                            for (index, arg) in arg_exprs.iter().enumerate() {
                                self.compile_expr(arg)?;
                                let arg_slot =
                                    self.define_local(&format!("__member_call_arg_{}", index));
                                self.emit_u16(Op::LOCAL_SET, arg_slot);
                                arg_slots.push(arg_slot);
                            }
                            self.emit_call_ref_with_arg_slots(
                                fn_tmp,
                                Some(receiver_slot),
                                &arg_slots,
                            );
                            self.emit_u16(Op::LOCAL_GET, receiver_slot);
                            self.emit(Op::REF_IS_NULL);
                            let line = self.line;
                            self.chunk().emit_if(line);
                            self.chunk().emit_else(line);
                            self.emit_u16(Op::LOCAL_GET, receiver_slot);
                            fn_call!(self, "wasm:js-undefined", "test", 1);
                            let line = self.line;
                            self.chunk().emit_if(line);
                            self.chunk().emit_else(line);
                            self.emit_fortran_member_receiver_writeback(object, obj_tmp);
                            self.chunk().emit_end(line);
                            self.chunk().emit_end(line);
                        }
                    } else {
                        let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                        for (index, arg) in arg_exprs.iter().enumerate() {
                            self.compile_expr(arg)?;
                            let arg_slot =
                                self.define_local(&format!("__member_call_arg_{}", index));
                            self.emit_u16(Op::LOCAL_SET, arg_slot);
                            arg_slots.push(arg_slot);
                        }
                        self.emit_call_ref_with_arg_slots(fn_tmp, Some(receiver_slot), &arg_slots);
                        self.emit_u16(Op::LOCAL_GET, receiver_slot);
                        self.emit(Op::REF_IS_NULL);
                        let line = self.line;
                        self.chunk().emit_if(line);
                        self.chunk().emit_else(line);
                        self.emit_u16(Op::LOCAL_GET, receiver_slot);
                        fn_call!(self, "wasm:js-undefined", "test", 1);
                        let line = self.line;
                        self.chunk().emit_if(line);
                        self.chunk().emit_else(line);
                        self.emit_fortran_member_receiver_writeback(object, obj_tmp);
                        self.chunk().emit_end(line);
                        self.chunk().emit_end(line);
                    }
                }
            }
            if let Some(line) = primitive_tostring_if {
                self.chunk().emit_end(line);
            }
            if let Some(line) = member_index_if {
                self.finish_member_index_call_path(callee, &arg_exprs, fn_tmp, line)?;
            }
            return Ok(());
        }

        // ── Simple call: name(args) / expr(args) ────────────────────
        if let ExprKind::Ident(name) = &callee.kind {
            if self.is_js_profile() && name == "Function" {
                for arg in &arg_exprs {
                    self.compile_expr(arg)?;
                }
                let idx = self.import("vybe:js", "function_constructor");
                self.emit_host_call(idx, arg_exprs.len() as u8);
                return Ok(());
            }

            if self.try_compile_fortran_derived_type_constructor(name, args)? {
                return Ok(());
            }

            let rest_signature = self
                .function_signatures
                .get(&self.canon(name))
                .and_then(|signatures| self.select_call_signature(signatures, args))
                .filter(|signature| signature.has_rest)
                .cloned();

            // ── ESM host-module import binding ──────────────────────────
            let key = self.canon(name);
            if let Some((module, func)) = self.host_import_bindings.get(&key).cloned() {
                let _ = (module, func);
                let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                for (index, arg) in arg_exprs.iter().enumerate() {
                    self.compile_expr_with_value_copy(arg)?;
                    let arg_slot = self.define_local(&format!("__host_import_call_arg_{}", index));
                    self.emit_u16(Op::LOCAL_SET, arg_slot);
                    arg_slots.push(arg_slot);
                }
                self.emit_var_get(name);
                let callee_slot = self.define_local("__host_import_call_callee");
                self.emit_u16(Op::LOCAL_SET, callee_slot);
                self.emit_call_ref_with_arg_slots(callee_slot, None, &arg_slots);
                return Ok(());
            }

            if self.is_php_profile()
                && (name.eq_ignore_ascii_case("exit") || name.eq_ignore_ascii_case("die"))
            {
                if let Some(arg) = arg_exprs.first() {
                    let arg_slot = self.define_local("__php_exit_arg");
                    self.compile_expr(arg)?;
                    self.emit_u16(Op::LOCAL_SET, arg_slot);

                    self.emit_u16(Op::LOCAL_GET, arg_slot);
                    let typeof_idx = self.import("ecma:value", "typeof");
                    self.emit_host_call(typeof_idx, 1);
                    self.emit_const(Value::String(Arc::from("string")));
                    {
                        let line = self.line;
                        crate::emitter::ops::emit_dyn_eq(self.chunk(), line);
                    };
                    let line = self.line;
                    crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                    self.chunk().emit_if(line);

                    let log_idx = self.import("wasi:logging/logging", "log");
                    let line = self.line;
                    self.emit_u16(Op::LOCAL_GET, arg_slot);
                    self.emit_common("php.echo_stringify", 1, line);
                    common::io::emit_print_with_import(self.chunk(), log_idx, 1, line);

                    self.chunk().emit_end(line);
                }

                self.emit(Op::NULL);
                self.emit_return_through_finally(1)?;
                return Ok(());
            }

            // Inside a class: bare call to a static method should bind to
            // the class object before any generic function lookup. Static
            // methods are also registered as ordinary functions, so this
            // must run ahead of `is_known_func`.
            if self.current_class.is_some()
                && (self.current_member_is_static || self.current_class_implicit_self)
            {
                let is_local = self.has_accessible_local_binding(name);
                if !is_local {
                    if let Some(class_name) = self.is_class_static_method(name) {
                        let cls_idx = self.str_const(&class_name);
                        self.emit_u16(Op::GLOBAL_GET, cls_idx);
                        let method_idx = self.str_const(&self.canon(name));
                        self.emit_u16(Op::STRUCT_GET, method_idx);
                        let fn_tmp = self.define_local("__bare_static_fn");
                        self.emit_u16(Op::LOCAL_SET, fn_tmp);

                        let method_canon = self.canon(name);
                        if let Some(param_modes) =
                            self.function_param_modes.get(&method_canon).cloned()
                        {
                            if param_modes
                                .iter()
                                .any(|mode| matches!(mode, PassBy::Ref | PassBy::Out))
                            {
                                let mut arg_slots = Vec::with_capacity(args.len());
                                for (index, arg) in args.iter().enumerate() {
                                    self.compile_ref_aware_call_arg(
                                        arg,
                                        param_modes.get(index).copied().unwrap_or(PassBy::Value),
                                    )?;
                                    let arg_slot = self
                                        .define_local(&format!("__bare_static_call_arg_{}", index));
                                    self.emit_u16(Op::LOCAL_SET, arg_slot);
                                    arg_slots.push(arg_slot);
                                }

                                self.emit_u16(Op::LOCAL_GET, fn_tmp);
                                for slot in &arg_slots {
                                    self.emit_u16(Op::LOCAL_GET, *slot);
                                }
                                self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);

                                let pack_slot = self.define_local("__bare_static_ref_call_pack");
                                self.emit_u16(Op::LOCAL_SET, pack_slot);
                                let mut ref_out_index = 1usize;
                                for (index, arg) in args.iter().enumerate() {
                                    if !matches!(
                                        param_modes.get(index),
                                        Some(PassBy::Ref | PassBy::Out)
                                    ) {
                                        continue;
                                    }
                                    self.emit_u16(Op::LOCAL_GET, pack_slot);
                                    self.emit_const(Value::F64(ref_out_index as f64));
                                    common::collections::emit_get(
                                        &mut self.chunks,
                                        self.current,
                                        self.line,
                                    );
                                    self.compile_assign_target(&arg.value)?;
                                    ref_out_index += 1;
                                }
                                self.emit_u16(Op::LOCAL_GET, pack_slot);
                                self.emit_const(Value::F64(0.0));
                                common::collections::emit_get(
                                    &mut self.chunks,
                                    self.current,
                                    self.line,
                                );
                                return Ok(());
                            }
                        }

                        if self.profile.name == "csharp" && args.len() == 1 && !args[0].spread {
                            if self
                                .resolve_static_method_overload_for_type(
                                    &class_name,
                                    name,
                                    &arg_exprs,
                                )
                                .is_some_and(|overload| overload.signature.has_rest)
                            {
                                self.emit_variadic_array_call_from_local(fn_tmp, &args[0].value)?;
                                return Ok(());
                            }
                        }

                        let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                        for (index, arg) in arg_exprs.iter().enumerate() {
                            self.compile_expr_with_value_copy(arg)?;
                            let arg_slot =
                                self.define_local(&format!("__bare_static_call_arg_{}", index));
                            self.emit_u16(Op::LOCAL_SET, arg_slot);
                            arg_slots.push(arg_slot);
                        }
                        self.emit_call_ref_with_arg_slots(fn_tmp, None, &arg_slots);
                        return Ok(());
                    }
                }
            }

            let fortran_arg_exprs: Vec<Expression> = arg_exprs.iter().cloned().cloned().collect();
            if !self.has_accessible_local_binding(name) {
                if let Some(target_name) =
                    self.resolve_fortran_interface_target(name, &fortran_arg_exprs)
                {
                    // Standalone `interface` blocks declare an external name
                    // (e.g. `area_circle`) whose overload target is the same
                    // symbol. Re-dispatching to that name loops forever; fall
                    // through to the normal external-function path instead.
                    if self.canon(&target_name) != self.canon(name) {
                        let callee = if let Some(module_name) =
                            self.enum_members.get(&self.canon(&target_name)).cloned()
                        {
                            Expression::new(ExprKind::Member {
                                object: Box::new(Expression::ident(&module_name)),
                                field: target_name.clone(),
                                null_safe: false,
                            })
                        } else {
                            Expression::ident(&target_name)
                        };
                        self.compile_call(&callee, args)?;
                        return Ok(());
                    }
                }
            }

            let canonical_name = self.canon(name);
            let is_known_module_static = self
                .enum_members
                .get(&canonical_name)
                .and_then(|module_name| self.pending_classes.get(module_name))
                .is_some_and(|pending| {
                    pending
                        .static_method_names
                        .iter()
                        .any(|member| member == &canonical_name)
                });
            let is_known_func = self.defined_functions.contains(name)
                || (!self.case_sensitive
                    && self
                        .defined_functions
                        .iter()
                        .any(|g| g.eq_ignore_ascii_case(name)))
                || is_known_module_static;
            if !is_known_func && self.try_compile_builtin(name, &arg_exprs)? {
                return Ok(());
            }

            // VB array access: `arr(idx)` when `arr` is a known data variable
            // (local OR top-level global from `Dim arr(5)`) and is NOT a
            // declared function or class. VB syntactically overloads `()` for
            // both calls and indexing — the disambiguator is whether the head
            // is a callable function or a value. We must exclude both
            // `defined_functions` and `defined_classes` from the "looks like
            // a variable" set, otherwise `GetResult()` (function call) and
            // `New Result()` (class) would be mis-identified as indexing.
            if !is_known_func
                && arg_exprs.len() == 1
                && self.profile.parens_for_index
                && !self.is_php_profile()
            {
                let canon_name = self.canon(name);
                let is_local = self.has_accessible_local_binding(name);
                let is_global_var = self.defined_globals.contains(&canon_name)
                    && !self.defined_classes.contains(&canon_name)
                    && !self.defined_functions.contains(&canon_name);
                let is_callable_typed = self
                    .lookup_var_type_hint(name)
                    .is_some_and(Self::is_callable_type_hint);
                let is_indexable_typed = matches!(
                    self.lookup_var_type_hint(name)
                        .map(Self::normalize_type_hint)
                        .as_deref(),
                    Some(type_hint) if type_hint.ends_with("()") && !Self::is_callable_type_hint(type_hint)
                ) || self.lookup_array_binding(name).is_some();
                if (is_local || is_global_var) && !is_callable_typed {
                    if is_indexable_typed {
                        self.emit_var_get(name);
                        self.compile_array_index_operand_for_owner(callee, arg_exprs[0])?;
                        {
                            let l = self.line;
                            common::collections::emit_get(&mut self.chunks, self.current, l);
                        }
                        return Ok(());
                    }

                    // Ambiguous `value(arg)` forms need a runtime split:
                    // procedure dummy arguments and other first-class callables
                    // must call, while plain arrays still index.
                    let callee_slot = self.define_local("__paren_ambig_callee");
                    self.emit_var_get(name);
                    self.emit_u16(Op::LOCAL_SET, callee_slot);

                    let table_idx_key = self.str_const("__table_idx");
                    self.emit_u16(Op::LOCAL_GET, callee_slot);
                    self.emit_u16(Op::STRUCT_GET, table_idx_key);
                    let table_idx_slot = self.define_local("__paren_ambig_table_idx");
                    self.emit_u16(Op::LOCAL_SET, table_idx_slot);

                    self.emit_u16(Op::LOCAL_GET, table_idx_slot);
                    self.emit(Op::REF_IS_NULL);
                    let line = self.line;
                    self.chunk().emit_if_value(line);
                    self.emit_u16(Op::LOCAL_GET, callee_slot);
                    self.compile_array_index_operand_for_owner(callee, arg_exprs[0])?;
                    {
                        let l = self.line;
                        common::collections::emit_get(&mut self.chunks, self.current, l);
                    }
                    self.chunk().emit_else(line);
                    self.emit_u16(Op::LOCAL_GET, table_idx_slot);
                    fn_call!(self, "wasm:js-undefined", "test", 1);
                    let line = self.line;
                    self.chunk().emit_if_value(line);
                    self.emit_u16(Op::LOCAL_GET, callee_slot);
                    self.compile_array_index_operand_for_owner(callee, arg_exprs[0])?;
                    {
                        let l = self.line;
                        common::collections::emit_get(&mut self.chunks, self.current, l);
                    }
                    self.chunk().emit_else(line);

                    let receiver_key = self.str_const("__vybe_method_receiver");
                    self.emit_u16(Op::LOCAL_GET, callee_slot);
                    self.emit_u16(Op::STRUCT_GET, receiver_key);
                    let receiver_slot = self.define_local("__paren_ambig_receiver");
                    self.emit_u16(Op::LOCAL_SET, receiver_slot);

                    self.compile_expr_with_value_copy(&arg_exprs[0])?;
                    let arg_slot = self.define_local("__paren_ambig_arg_0");
                    self.emit_u16(Op::LOCAL_SET, arg_slot);
                    self.emit_call_ref_with_arg_slots(
                        callee_slot,
                        Some(receiver_slot),
                        &[arg_slot],
                    );
                    self.chunk().emit_end(line);
                    self.chunk().emit_end(line);
                    return Ok(());
                }
            }

            // Inside a class: bare method call → Me.method(args)
            // If name isn't a local variable and we're inside a class body,
            // resolve as Me.name() (implicit self for method calls).
            if self.current_class.is_some()
                && self.current_class_implicit_self
                && !self.current_member_is_static
            {
                let is_local = self.has_accessible_local_binding(name);
                if !is_local && !is_known_func {
                    if self.emit_self_ref() {
                        // Me.name(args) → load Me, dup, struct_get(name).
                        // Real methods receive `this`/Self as arg0, but callable
                        // fields (Pascal procedure/function members) should be
                        // invoked as plain function values.
                        let field_name = self.canon(name);
                        let prop = self.str_const(&field_name);
                        inst!(self, core_wasm::dup);
                        self.emit_u16(Op::STRUCT_GET, prop);
                        let fn_tmp = self.define_local("__bare_fn");
                        self.emit_u16(Op::LOCAL_SET, fn_tmp);
                        let obj_tmp = self.define_local("__bare_obj");
                        self.emit_u16(Op::LOCAL_SET, obj_tmp);

                        let is_callable_field = self
                            .lookup_implicit_self_field_type_hint(name)
                            .is_some_and(Self::is_callable_type_hint);
                        if (self.profile.name == "pascal" && self.is_class_field(name))
                            || is_callable_field
                        {
                            let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                            for (index, arg) in arg_exprs.iter().enumerate() {
                                self.compile_expr_with_value_copy(arg)?;
                                let arg_slot =
                                    self.define_local(&format!("__bare_field_call_arg_{}", index));
                                self.emit_u16(Op::LOCAL_SET, arg_slot);
                                arg_slots.push(arg_slot);
                            }
                            self.emit_call_ref_with_arg_slots(fn_tmp, None, &arg_slots);
                            return Ok(());
                        }

                        let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                        for (index, arg) in arg_exprs.iter().enumerate() {
                            self.compile_expr_with_value_copy(arg)?;
                            let arg_slot =
                                self.define_local(&format!("__bare_method_call_arg_{}", index));
                            self.emit_u16(Op::LOCAL_SET, arg_slot);
                            arg_slots.push(arg_slot);
                        }
                        self.emit_call_ref_with_arg_slots(fn_tmp, Some(obj_tmp), &arg_slots);
                        return Ok(());
                    }
                }
            }

            let has_spread = args.iter().any(|a| a.spread);
            if has_spread {
                if let Some(signature) = rest_signature.as_ref() {
                    let callee_slot = self.define_local("__packed_rest_spread_callee");
                    self.emit_var_get(name);
                    self.emit_u16(Op::LOCAL_SET, callee_slot);
                    self.emit_known_rest_call_from_local(callee_slot, None, args, signature)?;
                    return Ok(());
                }

                if self.profile.name == "php" && args.len() == 1 && args[0].spread {
                    if let Some(signature) = self
                        .function_signatures
                        .get(&self.canon(name))
                        .and_then(|signatures| {
                            signatures.iter().find(|signature| {
                                !signature.has_rest && !signature.param_names.is_empty()
                            })
                        })
                        .cloned()
                    {
                        let spread_slot = self.define_local("__php_named_unpack");
                        self.compile_expr(&args[0].value)?;
                        self.emit_u16(Op::LOCAL_SET, spread_slot);

                        let probe_slot = self.define_local("__php_named_unpack_probe");
                        self.emit_u16(Op::LOCAL_GET, spread_slot);
                        self.emit_const(Value::String(Arc::from(
                            signature.param_names[0].as_str(),
                        )));
                        common::collections::emit_get(&mut self.chunks, self.current, self.line);
                        self.emit_u16(Op::LOCAL_SET, probe_slot);

                        self.emit_u16(Op::LOCAL_GET, probe_slot);
                        self.emit(Op::REF_IS_NULL);
                        let line = self.line;
                        self.chunk().emit_if_value(line);

                        let line = self.line;
                        let args_slot = self.define_local("__spread_args");
                        common::collections::emit_array_new(
                            &mut self.chunks,
                            self.current,
                            0,
                            line,
                        );
                        self.emit_u16(Op::LOCAL_SET, args_slot);
                        let mut known_len: Option<usize> = Some(0);
                        for a in args {
                            if a.spread {
                                self.emit_u16(Op::LOCAL_GET, args_slot);
                                self.compile_expr(&a.value)?;
                                common::collections::emit_concat(
                                    &mut self.chunks,
                                    self.current,
                                    line,
                                );
                                self.emit_u16(Op::LOCAL_SET, args_slot);
                                if let ExprKind::Array(elems) = &a.value.kind {
                                    if let Some(ref mut k) = known_len {
                                        *k += elems.len();
                                    }
                                } else {
                                    known_len = None;
                                }
                            } else {
                                self.emit_u16(Op::LOCAL_GET, args_slot);
                                self.compile_expr(&a.value)?;
                                common::collections::emit_push(
                                    &mut self.chunks,
                                    self.current,
                                    line,
                                );
                                self.emit(Op::DROP);
                                if let Some(ref mut k) = known_len {
                                    *k += 1;
                                }
                            }
                        }
                        self.emit_var_get(name);
                        inst!(self, core_wasm::undefined);
                        self.emit_u16(Op::LOCAL_GET, args_slot);
                        fn_call!(self, "ecma:function", "apply", 3);

                        self.chunk().emit_else(line);

                        self.emit_var_get(name);
                        for param_name in &signature.param_names {
                            self.emit_u16(Op::LOCAL_GET, spread_slot);
                            self.emit_const(Value::String(Arc::from(param_name.as_str())));
                            common::collections::emit_get(
                                &mut self.chunks,
                                self.current,
                                self.line,
                            );
                        }
                        self.emit_u8(Op::CALL_REF, signature.param_names.len() as u8);
                        self.chunk().emit_end(line);
                        return Ok(());
                    }
                }

                // Spread args: build a flat args array, then spread onto
                // stack and call. Stash the accumulator in a local so
                // `ecma:array.push` (returns new length per
                // ECMA-262) and `ecma:array.concat` (returns new
                // array) can both drive the same pattern.
                let line = self.line;
                let args_slot = self.define_local("__spread_args");
                common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
                self.emit_u16(Op::LOCAL_SET, args_slot);
                let mut known_len: Option<usize> = Some(0);
                for a in args {
                    if a.spread {
                        // new_arr = concat(args, spread)
                        self.emit_u16(Op::LOCAL_GET, args_slot);
                        self.compile_expr(&a.value)?;
                        common::collections::emit_concat(&mut self.chunks, self.current, line);
                        self.emit_u16(Op::LOCAL_SET, args_slot);
                        if let ExprKind::Array(elems) = &a.value.kind {
                            if let Some(ref mut k) = known_len {
                                *k += elems.len();
                            }
                        } else {
                            known_len = None;
                        }
                    } else {
                        self.emit_u16(Op::LOCAL_GET, args_slot);
                        self.compile_expr(&a.value)?;
                        common::collections::emit_push(&mut self.chunks, self.current, line);
                        self.emit(Op::DROP); // drop new_length returned by push
                        if let Some(ref mut k) = known_len {
                            *k += 1;
                        }
                    }
                }
                let is_local = self.scope().resolve(name).is_some()
                    || (!self.case_sensitive && self.scope().resolve_ci(name).is_some())
                    || self.has_static_local_binding(name);
                let canon_name = self.canon(name);
                let is_direct_global = self.defined_globals.contains(&canon_name)
                    || self.defined_functions.contains(&canon_name);
                if !is_local {
                    if is_direct_global {
                        self.emit_var_get(name);
                    } else if let Some(module_name) = self.enum_members.get(&canon_name).cloned() {
                        let prefers_direct_module_global = self
                            .pending_classes
                            .get(&module_name)
                            .is_some_and(|pending| {
                                pending
                                    .static_method_names
                                    .iter()
                                    .any(|member| member == &canon_name)
                            });
                        if prefers_direct_module_global {
                            self.emit_var_get(name);
                        } else {
                            let module_idx = self.str_const(&module_name);
                            self.emit_u16(Op::GLOBAL_GET, module_idx);
                            let member_idx = self.str_const(&canon_name);
                            self.emit_u16(Op::STRUCT_GET, member_idx);
                        }
                    } else {
                        self.emit_var_get(name);
                    }
                } else {
                    self.emit_var_get(name);
                }
                let callee_slot = self.define_local("__ident_spread_callee");
                self.emit_u16(Op::LOCAL_SET, callee_slot);
                self.emit_php_dynamic_function_name_resolution(callee_slot);
                let receiver_key = self.str_const("__vybe_method_receiver");
                self.emit_u16(Op::LOCAL_GET, callee_slot);
                self.emit_u16(Op::STRUCT_GET, receiver_key);
                let receiver_slot = self.define_local("__ident_spread_receiver");
                self.emit_u16(Op::LOCAL_SET, receiver_slot);
                self.emit_call_ref_with_args_array(
                    callee_slot,
                    Some(receiver_slot),
                    args_slot,
                    known_len,
                );
                return Ok(());
            }
            if self.is_python_profile() && !is_known_func {
                let callee_slot = self.define_local("__py_call_target");
                self.emit_var_get(name);
                self.emit_u16(Op::LOCAL_SET, callee_slot);

                self.emit_u16(Op::LOCAL_GET, callee_slot);
                let typeof_idx = self.import("ecma:value", "typeof");
                self.emit_host_call(typeof_idx, 1);
                self.emit_const(Value::String(Arc::from("function")));
                {
                    let line = self.line;
                    crate::emitter::ops::emit_dyn_eq(self.chunk(), line);
                };
                let line = self.line;
                crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                self.chunk().emit_if_value(line);

                self.emit_u16(Op::LOCAL_GET, callee_slot);
                for a in &arg_exprs {
                    self.compile_expr(a)?;
                }
                self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);

                self.chunk().emit_else(line);
                self.emit_u16(Op::LOCAL_GET, callee_slot);
                let call_prop = self.str_const("call");
                self.emit_u16(Op::STRUCT_GET, call_prop);
                let call_slot = self.define_local("__py_call_method");
                self.emit_u16(Op::LOCAL_SET, call_slot);
                self.emit_u16(Op::LOCAL_GET, call_slot);
                self.emit(Op::REF_IS_NULL);
                let line = self.line;
                self.chunk().emit_if_value(line);

                self.emit_u16(Op::LOCAL_GET, callee_slot);
                let dunder_prop = self.str_const("__call__");
                self.emit_u16(Op::STRUCT_GET, dunder_prop);
                let dunder_slot = self.define_local("__py_dunder_call_method");
                self.emit_u16(Op::LOCAL_SET, dunder_slot);
                self.emit_u16(Op::LOCAL_GET, dunder_slot);
                self.emit(Op::REF_IS_NULL);
                let line = self.line;
                self.chunk().emit_if_value(line);
                inst!(self, core_wasm::undefined);
                self.chunk().emit_else(line);
                self.emit_u16(Op::LOCAL_GET, dunder_slot);
                self.emit_u16(Op::LOCAL_GET, callee_slot);
                for a in &arg_exprs {
                    self.compile_expr(a)?;
                }
                self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                self.chunk().emit_end(line);

                self.chunk().emit_else(line);
                self.emit_u16(Op::LOCAL_GET, call_slot);
                self.emit_u16(Op::LOCAL_GET, callee_slot);
                for a in &arg_exprs {
                    self.compile_expr(a)?;
                }
                self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                self.chunk().emit_end(line);
                self.chunk().emit_end(line);
                return Ok(());
            }

            let callee_slot = self.define_local("__direct_call_callee");
            let is_local = self.scope().resolve(name).is_some()
                || (!self.case_sensitive && self.scope().resolve_ci(name).is_some())
                || self.has_static_local_binding(name);
            let canon_name = self.canon(name);
            let is_direct_global = self.defined_globals.contains(&canon_name)
                || self.defined_functions.contains(&canon_name);
            if !is_local {
                if is_direct_global {
                    self.emit_var_get(name);
                } else if let Some(module_name) = self.enum_members.get(&canon_name).cloned() {
                    let prefers_direct_module_global = self
                        .pending_classes
                        .get(&module_name)
                        .is_some_and(|pending| {
                            pending
                                .static_method_names
                                .iter()
                                .any(|member| member == &canon_name)
                        });
                    if prefers_direct_module_global {
                        self.emit_var_get(name);
                    } else {
                        let module_idx = self.str_const(&module_name);
                        self.emit_u16(Op::GLOBAL_GET, module_idx);
                        let member_idx = self.str_const(&canon_name);
                        self.emit_u16(Op::STRUCT_GET, member_idx);
                    }
                } else {
                    self.emit_var_get(name);
                }
            } else {
                self.emit_var_get(name);
            }
            self.emit_u16(Op::LOCAL_SET, callee_slot);
            self.emit_php_dynamic_function_name_resolution(callee_slot);
            let receiver_key = self.str_const("__vybe_method_receiver");
            self.emit_u16(Op::LOCAL_GET, callee_slot);
            self.emit_u16(Op::STRUCT_GET, receiver_key);
            let receiver_slot = self.define_local("__direct_call_receiver");
            self.emit_u16(Op::LOCAL_SET, receiver_slot);
            if let Some(signature) = rest_signature.as_ref() {
                let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                for (index, arg) in arg_exprs.iter().enumerate() {
                    self.compile_expr_with_value_copy(arg)?;
                    let arg_slot = self.define_local(&format!("__direct_rest_call_arg_{}", index));
                    self.emit_u16(Op::LOCAL_SET, arg_slot);
                    arg_slots.push(arg_slot);
                }

                let _ = signature;
                self.emit_call_ref_with_arg_slots(callee_slot, Some(receiver_slot), &arg_slots);
                return Ok(());
            }
            if let Some(param_modes) = self.function_param_modes.get(&self.canon(name)).cloned() {
                let needs_ref_aware_args = param_modes
                    .iter()
                    .copied()
                    .any(|mode| self.mode_needs_ref_aware_call_handling(mode));
                let needs_packed_result = param_modes
                    .iter()
                    .copied()
                    .any(|mode| self.mode_needs_call_writeback(mode));
                if needs_ref_aware_args {
                    let mut arg_slots = Vec::with_capacity(args.len());
                    for (index, arg) in args.iter().enumerate() {
                        self.compile_ref_aware_call_arg(
                            arg,
                            param_modes.get(index).copied().unwrap_or(PassBy::Value),
                        )?;

                        let arg_slot = self.define_local(&format!("__direct_call_arg_{}", index));
                        self.emit_u16(Op::LOCAL_SET, arg_slot);
                        arg_slots.push(arg_slot);
                    }

                    self.emit_call_ref_with_arg_slots(callee_slot, Some(receiver_slot), &arg_slots);

                    if !needs_packed_result {
                        return Ok(());
                    }

                    let pack_slot = self.define_local("__ref_call_pack");
                    self.emit_u16(Op::LOCAL_SET, pack_slot);
                    let mut ref_out_index = 1usize;
                    for (index, arg) in args.iter().enumerate() {
                        if !param_modes
                            .get(index)
                            .copied()
                            .is_some_and(|mode| self.mode_needs_call_writeback(mode))
                        {
                            continue;
                        }
                        self.emit_u16(Op::LOCAL_GET, pack_slot);
                        self.emit_const(Value::F64(ref_out_index as f64));
                        common::collections::emit_get(&mut self.chunks, self.current, self.line);
                        self.compile_assign_target(&arg.value)?;
                        ref_out_index += 1;
                    }
                    self.emit_u16(Op::LOCAL_GET, pack_slot);
                    self.emit_const(Value::F64(0.0));
                    common::collections::emit_get(&mut self.chunks, self.current, self.line);
                    return Ok(());
                }
            }

            if args.iter().any(|arg| arg.by_ref) {
                let mut arg_slots = Vec::with_capacity(args.len());
                for (index, arg) in args.iter().enumerate() {
                    if arg.by_ref {
                        self.compile_expr(&arg.value)?;
                    } else {
                        self.compile_expr_with_value_copy(&arg.value)?;
                    }
                    let arg_slot = self.define_local(&format!("__direct_call_arg_{}", index));
                    self.emit_u16(Op::LOCAL_SET, arg_slot);
                    arg_slots.push(arg_slot);
                }

                self.emit_call_ref_with_arg_slots(callee_slot, Some(receiver_slot), &arg_slots);

                let result_slot = self.define_local("__direct_call_result");
                self.emit_u16(Op::LOCAL_SET, result_slot);

                let mut ref_out_index = 1usize;
                for arg in args {
                    if !arg.by_ref {
                        continue;
                    }
                    self.emit_u16(Op::LOCAL_GET, result_slot);
                    self.emit_const(Value::F64(ref_out_index as f64));
                    common::collections::emit_get(&mut self.chunks, self.current, self.line);
                    self.compile_assign_target(&arg.value)?;
                    ref_out_index += 1;
                }
                self.emit_u16(Op::LOCAL_GET, result_slot);
                self.emit_const(Value::F64(0.0));
                common::collections::emit_get(&mut self.chunks, self.current, self.line);
                return Ok(());
            }

            let mut arg_slots = Vec::with_capacity(arg_exprs.len());
            for (index, arg) in arg_exprs.iter().enumerate() {
                self.compile_expr_with_value_copy(arg)?;
                let arg_slot = self.define_local(&format!("__direct_call_arg_{}", index));
                self.emit_u16(Op::LOCAL_SET, arg_slot);
                arg_slots.push(arg_slot);
            }

            // For fixed-arity generator functions, pad missing optional args
            // with `undefined` so the GEN_NEXT null resume value never lands
            // in an optional-parameter slot and prevents default application.
            // (GEN_NEXT appends `null` as the last element of the arg list it
            // builds from __bound_args, so without padding it would be placed
            // in the first missing slot rather than the control slot.)
            if let Some(&gen_param_count) = self.generator_param_counts.get(&self.canon(name)) {
                let provided = arg_slots.len();
                for index in provided..gen_param_count {
                    let pad_slot = self.define_local(&format!("__gen_pad_arg_{}", index));
                    inst!(self, core_wasm::undefined);
                    self.emit_u16(Op::LOCAL_SET, pad_slot);
                    arg_slots.push(pad_slot);
                }
            }

            self.emit_call_ref_with_arg_slots(callee_slot, Some(receiver_slot), &arg_slots);
            return Ok(());
        }

        // ── Computed-member call: `obj[key](args)` ───────────────────
        // For JS profile, treat this like a method call so `__js_this`
        // is bound to `obj` before invocation. Without this binding the
        // callee body sees a stale __js_this and `this.x` traps. Same
        // semantics as ECMA-262 §13.3.7 (CallMemberExpression).
        if self.is_js_profile() {
            if let ExprKind::Index { object, index, .. } = &callee.kind {
                if arg_exprs.is_empty()
                    && matches!(&object.kind, ExprKind::Array(_))
                    && matches!(
                        &index.kind,
                        ExprKind::Member { object, field, null_safe: false }
                            if matches!(&object.kind, ExprKind::Ident(name) if name == "Symbol")
                                && field == "iterator"
                    )
                {
                    self.compile_expr(object)?;
                    let values_idx = self.import("ecma:array", "values");
                    self.emit_host_call(values_idx, 1);
                    return Ok(());
                }
                let obj_tmp = self.define_local("__js_idx_obj");
                self.compile_expr(object)?;
                self.emit_u16(Op::LOCAL_SET, obj_tmp);
                let saved_js_this = self.save_js_this("__js_prev_this_idx");
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.set_js_this_from_stack();
                let key_tmp = self.define_local("__js_idx_key");
                match &index.kind {
                    ExprKind::Member {
                        object,
                        field,
                        null_safe: false,
                    } if matches!(&object.kind, ExprKind::Ident(name) if name == "Symbol") => {
                        let fallback_key = match field.as_str() {
                            "iterator" => Some("iterator"),
                            "asyncIterator" => Some("asyncIterator"),
                            "toPrimitive" => Some("toprimitive"),
                            "hasInstance" => Some("hasinstance"),
                            _ => None,
                        };
                        if let Some(fallback_key) = fallback_key {
                            self.emit_const(Value::String(Arc::from(fallback_key)));
                        } else {
                            self.compile_expr(index)?;
                        }
                    }
                    _ => self.compile_expr(index)?,
                }
                self.emit_u16(Op::LOCAL_SET, key_tmp);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::LOCAL_GET, key_tmp);
                let line = self.line;
                common::collections::emit_get(&mut self.chunks, self.current, line);
                let callee_tmp = self.define_local("__js_idx_callee");
                self.emit_u16(Op::LOCAL_SET, callee_tmp);

                self.emit_u16(Op::LOCAL_GET, callee_tmp);
                self.emit(Op::REF_IS_NULL);
                let lookup = self.str_const("__vybe_js_get_method");
                let line = self.line;
                self.chunk().emit_if(line);
                self.emit_u16(Op::GLOBAL_GET, lookup);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                match &index.kind {
                    ExprKind::Member {
                        object,
                        field,
                        null_safe: false,
                    } if matches!(&object.kind, ExprKind::Ident(name) if name == "Symbol") => {
                        let fallback_key = match field.as_str() {
                            "iterator" => Some("iterator"),
                            "asyncIterator" => Some("asyncIterator"),
                            "toPrimitive" => Some("toprimitive"),
                            "hasInstance" => Some("hasinstance"),
                            _ => None,
                        };
                        if let Some(fallback_key) = fallback_key {
                            self.emit_const(Value::String(Arc::from(fallback_key)));
                        } else {
                            self.emit_u16(Op::LOCAL_GET, key_tmp);
                        }
                    }
                    _ => self.emit_u16(Op::LOCAL_GET, key_tmp),
                }
                self.emit_u8(Op::CALL_REF, 2);
                self.emit_u16(Op::LOCAL_SET, callee_tmp);
                self.chunk().emit_else(line);
                self.emit_u16(Op::LOCAL_GET, callee_tmp);
                fn_call!(self, "wasm:js-undefined", "test", 1);
                let line = self.line;
                self.chunk().emit_if(line);
                self.emit_u16(Op::GLOBAL_GET, lookup);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                match &index.kind {
                    ExprKind::Member {
                        object,
                        field,
                        null_safe: false,
                    } if matches!(&object.kind, ExprKind::Ident(name) if name == "Symbol") => {
                        let fallback_key = match field.as_str() {
                            "iterator" => Some("iterator"),
                            "asyncIterator" => Some("asyncIterator"),
                            "toPrimitive" => Some("toprimitive"),
                            "hasInstance" => Some("hasinstance"),
                            _ => None,
                        };
                        if let Some(fallback_key) = fallback_key {
                            self.emit_const(Value::String(Arc::from(fallback_key)));
                        } else {
                            self.emit_u16(Op::LOCAL_GET, key_tmp);
                        }
                    }
                    _ => self.emit_u16(Op::LOCAL_GET, key_tmp),
                }
                self.emit_u8(Op::CALL_REF, 2);
                self.emit_u16(Op::LOCAL_SET, callee_tmp);
                self.chunk().emit_end(line);
                self.chunk().emit_end(line);

                self.emit_u16(Op::LOCAL_GET, callee_tmp);
                for a in &arg_exprs {
                    self.compile_expr(a)?;
                }
                self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                let result_slot = self.define_local("__js_idx_result");
                self.emit_u16(Op::LOCAL_SET, result_slot);
                self.restore_js_this(saved_js_this);
                self.emit_u16(Op::LOCAL_GET, result_slot);
                return Ok(());
            }
        }

        if self.profile.parens_for_index && !arg_exprs.is_empty() {
            let is_bound_array = matches!(&callee.kind,
                ExprKind::Ident(name) if self.lookup_array_binding(name).is_some()
            );
            let is_indexable_typed = is_bound_array
                || self
                    .infer_expr_type_hint(callee)
                    .as_deref()
                    .map(Self::normalize_type_hint)
                    .is_some_and(|type_hint| {
                        type_hint.ends_with("()") && !Self::is_callable_type_hint(&type_hint)
                    });
            if is_indexable_typed {
                self.compile_expr(callee)?;
                for arg in &arg_exprs {
                    self.compile_array_index_operand_for_owner(callee, arg)?;
                    let line = self.line;
                    common::collections::emit_get(&mut self.chunks, self.current, line);
                }
                return Ok(());
            }
        }

        // ── Fallback: general expression call ───────────────────────
        self.compile_expr(callee)?;
        let callee_slot = self.define_local("__call_ref_callee");
        self.emit_u16(Op::LOCAL_SET, callee_slot);
        self.emit_php_dynamic_function_name_resolution(callee_slot);

        let result_slot = self.define_local("__call_ref_result");
        self.emit(Op::NULL);
        self.emit_u16(Op::LOCAL_SET, result_slot);
        let runtime_index_matched = self.define_local("__call_ref_runtime_index_matched");
        self.emit_const(Value::I32(0));
        self.emit_u16(Op::LOCAL_SET, runtime_index_matched);

        if self.profile.parens_for_index
            && !arg_exprs.is_empty()
            && matches!(&callee.kind, ExprKind::Call { .. } | ExprKind::Index { .. })
        {
            self.emit_u16(Op::LOCAL_GET, callee_slot);
            fn_call!(self, "ecma:array", "isArray", 1);
            let line = self.line;
            self.chunk().emit_if(line);

            self.emit_u16(Op::LOCAL_GET, callee_slot);
            for arg in &arg_exprs {
                self.compile_array_index_operand_for_owner(callee, arg)?;
                let line = self.line;
                common::collections::emit_get(&mut self.chunks, self.current, line);
            }
            self.emit_u16(Op::LOCAL_SET, result_slot);
            self.emit_const(Value::I32(1));
            self.emit_u16(Op::LOCAL_SET, runtime_index_matched);
            self.chunk().emit_end(line);
        }

        self.emit_u16(Op::LOCAL_GET, runtime_index_matched);
        self.emit(Op::I32_EQZ);
        let line = self.line;
        self.chunk().emit_if(line);

        let has_by_ref_args = args.iter().any(|arg| arg.by_ref);
        let receiver_key = self.str_const("__vybe_method_receiver");
        self.emit_u16(Op::LOCAL_GET, callee_slot);
        self.emit_u16(Op::STRUCT_GET, receiver_key);
        let receiver_slot = self.define_local("__call_ref_receiver");
        self.emit_u16(Op::LOCAL_SET, receiver_slot);

        let mut arg_slots = Vec::with_capacity(arg_exprs.len());
        for (index, arg) in arg_exprs.iter().enumerate() {
            self.compile_expr(arg)?;
            let arg_slot = self.define_local(&format!("__call_ref_arg_{}", index));
            self.emit_u16(Op::LOCAL_SET, arg_slot);
            arg_slots.push(arg_slot);
        }

        // Route through the shared dispatcher so a callee with a rest
        // parameter (e.g. a returned `(...more) => …`) gets its trailing
        // args packed via the runtime `__vybe_rest_fixed_arity` stamp —
        // a plain `CALL_REF` here would pass them unpacked and the rest
        // array would never form. Handles receiver-null/undefined branching
        // and `this` binding internally; leaves the result on the stack.
        let saved_js_new_target = self.save_js_new_target("__js_prev_new_target_call_ref");
        self.set_js_new_target_undefined();
        self.emit_call_ref_with_arg_slots(callee_slot, Some(receiver_slot), &arg_slots);
        self.emit_u16(Op::LOCAL_SET, result_slot);
        self.restore_js_new_target(saved_js_new_target);
        self.chunk().emit_end(line); // close the runtime_index_matched `if`

        if has_by_ref_args {
            let mut ref_out_index = 1usize;
            for arg in args {
                if !arg.by_ref {
                    continue;
                }
                self.emit_u16(Op::LOCAL_GET, result_slot);
                self.emit_const(Value::F64(ref_out_index as f64));
                common::collections::emit_get(&mut self.chunks, self.current, self.line);
                self.compile_assign_target(&arg.value)?;
                ref_out_index += 1;
            }
            self.emit_u16(Op::LOCAL_GET, result_slot);
            self.emit_const(Value::F64(0.0));
            common::collections::emit_get(&mut self.chunks, self.current, self.line);
        } else {
            self.emit_u16(Op::LOCAL_GET, result_slot);
        }

        Ok(())
    }

    fn try_compile_dotnet_guid_try_parse(
        &mut self,
        callee: &Expression,
        args: &[Argument],
    ) -> Result<bool, String> {
        if args.len() != 2 {
            return Ok(false);
        }
        let is_guid_try_parse = match &callee.kind {
            ExprKind::Member { object, field, .. } if field.eq_ignore_ascii_case("TryParse") => {
                terminal_type_name(object)
                    .is_some_and(|type_name| type_name.eq_ignore_ascii_case("Guid"))
            }
            _ => false,
        };
        if !is_guid_try_parse {
            return Ok(false);
        }

        let line = self.line;
        self.compile_expr(&args[0].value)?;
        self.emit_common("dotnet.guid_try_parse", 1, line);

        let parsed_slot = self.define_local("__guid_try_parse_value");
        self.emit_u16(Op::LOCAL_SET, parsed_slot);

        self.emit_u16(Op::LOCAL_GET, parsed_slot);
        self.emit(Op::REF_IS_NULL);
        let line = self.line;
        self.chunk().emit_if_value(line);

        self.emit(Op::NULL);
        self.compile_assign_target(&args[1].value)?;
        inst!(self, core_wasm::bool_const, false);

        self.chunk().emit_else(line);

        self.emit_u16(Op::LOCAL_GET, parsed_slot);
        self.compile_assign_target(&args[1].value)?;
        if let ExprKind::Ident(name) = &args[1].value.kind {
            let normalized = Self::normalize_type_hint("Guid");
            if let Some(slot) = self.scope().resolve_ci(name) {
                if let Some(local) = self
                    .scope_mut()
                    .locals
                    .iter_mut()
                    .rev()
                    .find(|local| local.slot == slot)
                {
                    local.type_hint = Some(normalized.clone());
                }
            } else {
                self.global_type_hints
                    .insert(self.canon(name), normalized.clone());
            }
        }
        inst!(self, core_wasm::bool_const, true);
        self.chunk().emit_end(line);
        Ok(true)
    }

    fn try_compile_dotnet_numeric_try_parse(
        &mut self,
        callee: &Expression,
        args: &[Argument],
    ) -> Result<bool, String> {
        if args.len() != 2 {
            return Ok(false);
        }
        let parsed_type = match &callee.kind {
            ExprKind::Member { object, field, .. } if field.eq_ignore_ascii_case("TryParse") => {
                terminal_type_name(object)
            }
            _ => None,
        };
        let Some(type_name) = parsed_type else {
            return Ok(false);
        };
        let normalized = Self::normalize_type_hint(&type_name);
        if normalized != "int" && normalized != "int32" {
            return Ok(false);
        }

        self.compile_expr(&args[0].value)?;
        let number_idx = self.import("ecma:number", "Number");
        self.emit_host_call(number_idx, 1);

        let parsed_slot = self.define_local("__numeric_try_parse_value");
        self.emit_u16(Op::LOCAL_SET, parsed_slot);

        self.emit_u16(Op::LOCAL_GET, parsed_slot);
        self.emit_u16(Op::LOCAL_GET, parsed_slot);
        {
            let line = self.line;
            crate::emitter::ops::emit_dyn_eq(self.chunk(), line);
        };
        let line = self.line;
        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if_value(line);

        self.emit_u16(Op::LOCAL_GET, parsed_slot);
        self.emit(Op::F64_FLOOR);
        self.compile_assign_target(&args[1].value)?;
        if let ExprKind::Ident(name) = &args[1].value.kind {
            let normalized = Self::normalize_type_hint("int");
            if let Some(slot) = self.scope().resolve_ci(name) {
                if let Some(local) = self
                    .scope_mut()
                    .locals
                    .iter_mut()
                    .rev()
                    .find(|local| local.slot == slot)
                {
                    local.type_hint = Some(normalized.clone());
                }
            } else {
                self.global_type_hints.insert(self.canon(name), normalized);
            }
        }
        inst!(self, core_wasm::bool_const, true);
        self.chunk().emit_else(line);
        self.emit_const(Value::F64(0.0));
        self.compile_assign_target(&args[1].value)?;
        inst!(self, core_wasm::bool_const, false);
        self.chunk().emit_end(line);
        Ok(true)
    }

    fn try_compile_dotnet_dictionary_try_get_value(
        &mut self,
        callee: &Expression,
        args: &[Argument],
    ) -> Result<bool, String> {
        if args.len() != 2 {
            return Ok(false);
        }
        let ExprKind::Member { object, field, .. } = &callee.kind else {
            return Ok(false);
        };
        if !field.eq_ignore_ascii_case("TryGetValue") {
            return Ok(false);
        }
        let is_dictionary = resolve_receiver_type_hint(self, object)
            .as_deref()
            .map(Self::is_dictionary_type_hint)
            .unwrap_or(false);
        if !is_dictionary {
            return Ok(false);
        }

        if let ExprKind::Ident(name) = &args[1].value.kind {
            let unresolved = self.scope().resolve(name).is_none()
                && (!self.case_sensitive && self.scope().resolve_ci(name).is_none()
                    || self.case_sensitive)
                && !self.defined_globals.contains(&self.canon(name));
            if unresolved {
                self.define_local_typed(name, None);
            }
        }

        self.compile_expr(object)?;
        let map_slot = self.define_local("__dict_try_get_map");
        self.emit_u16(Op::LOCAL_SET, map_slot);

        self.compile_expr(&args[0].value)?;
        if self.expr_uses_case_insensitive_string_keys(object) {
            let line = self.line;
            common::strings::emit_to_lower(self.chunk(), line);
        }
        let key_slot = self.define_local("__dict_try_get_key");
        self.emit_u16(Op::LOCAL_SET, key_slot);

        let has_idx = self.import("ecma:map", "has");
        self.emit_u16(Op::LOCAL_GET, map_slot);
        self.emit_u16(Op::LOCAL_GET, key_slot);
        self.emit_host_call(has_idx, 2);
        let has_slot = self.define_local("__dict_try_get_has");
        self.emit_u16(Op::LOCAL_SET, has_slot);

        self.emit_u16(Op::LOCAL_GET, has_slot);
        let line = self.line;
        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if_value(line);

        self.emit_u16(Op::LOCAL_GET, map_slot);
        let getter_key = self.str_const("__get___index__");
        self.emit_u16(Op::STRUCT_GET, getter_key);
        let getter_slot = self.define_local("__dict_try_get_getter");
        self.emit_u16(Op::LOCAL_SET, getter_slot);

        self.emit_u16(Op::LOCAL_GET, getter_slot);
        self.emit(Op::REF_IS_NULL);
        let line = self.line;
        self.chunk().emit_if_value(line);

        self.emit_u16(Op::LOCAL_GET, map_slot);
        self.emit_u16(Op::LOCAL_GET, key_slot);
        common::collections::emit_get(&mut self.chunks, self.current, self.line);

        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, getter_slot);
        self.emit_u16(Op::LOCAL_GET, map_slot);
        self.emit_u16(Op::LOCAL_GET, key_slot);
        self.emit_u8(Op::CALL_REF, 2);
        self.chunk().emit_end(line);

        self.compile_assign_target(&args[1].value)?;
        inst!(self, core_wasm::bool_const, true);

        self.chunk().emit_else(line);
        self.emit(Op::NULL);
        self.compile_assign_target(&args[1].value)?;
        inst!(self, core_wasm::bool_const, false);
        self.chunk().emit_end(line);
        Ok(true)
    }

    fn try_compile_dotnet_case_insensitive_collection_call(
        &mut self,
        callee: &Expression,
        args: &[Argument],
    ) -> Result<bool, String> {
        let ExprKind::Member { object, field, .. } = &callee.kind else {
            return Ok(false);
        };
        if !self.expr_uses_case_insensitive_string_keys(object) {
            return Ok(false);
        }

        let receiver_type = resolve_receiver_type_hint(self, object).unwrap_or_default();
        let normalized = Self::normalize_type_hint(&receiver_type);
        let line = self.line;

        if Self::is_dictionary_type_hint(&normalized) {
            match (field.as_str(), args.len()) {
                ("Add", 2) => {
                    let obj_slot = self.define_local("__dict_add_obj");
                    let key_slot = self.define_local("__dict_add_key");
                    let keys_slot = self.define_local("__dict_add_keys");

                    self.compile_expr(object)?;
                    self.emit_u16(Op::LOCAL_SET, obj_slot);

                    self.compile_collection_key(object, &args[0].value)?;
                    self.emit_u16(Op::LOCAL_SET, key_slot);

                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit_u16(Op::LOCAL_GET, key_slot);
                    self.compile_expr(&args[1].value)?;
                    let idx = self.import("ecma:map", "set");
                    self.emit_host_call(idx, 3);
                    self.emit(Op::DROP);

                    let keys_key = self.str_const("__keys");
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit_u16(Op::STRUCT_GET, keys_key);
                    self.emit_u16(Op::LOCAL_SET, keys_slot);

                    self.emit_u16(Op::LOCAL_GET, keys_slot);
                    self.emit(Op::REF_IS_NULL);
                    let line = self.line;
                    self.chunk().emit_if(line);

                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
                    inst!(self, core_wasm::dup);
                    self.emit_u16(Op::LOCAL_SET, keys_slot);
                    self.emit_u16(Op::STRUCT_SET, keys_key);
                    self.emit(Op::DROP);

                    self.chunk().emit_end(line);
                    self.emit_u16(Op::LOCAL_GET, keys_slot);
                    self.emit_u16(Op::LOCAL_GET, key_slot);
                    common::collections::emit_push(&mut self.chunks, self.current, line);
                    self.emit(Op::DROP);
                    return Ok(true);
                }
                ("ContainsKey", 1) => {
                    self.compile_expr(object)?;
                    self.compile_collection_key(object, &args[0].value)?;
                    let idx = self.import("ecma:map", "has");
                    self.emit_host_call(idx, 2);
                    return Ok(true);
                }
                ("Remove", 1) => {
                    self.compile_expr(object)?;
                    self.compile_collection_key(object, &args[0].value)?;
                    let idx = self.import("ecma:map", "delete");
                    self.emit_host_call(idx, 2);
                    return Ok(true);
                }
                _ => {}
            }
        }

        if normalized.contains("hashset") || normalized.contains("sortedset") {
            match (field.as_str(), args.len()) {
                ("Add", 1) => {
                    self.compile_expr(object)?;
                    self.compile_collection_key(object, &args[0].value)?;
                    self.emit_common("dotnet.hashset_add", 2, line);
                    return Ok(true);
                }
                ("Contains", 1) => {
                    self.compile_expr(object)?;
                    self.compile_collection_key(object, &args[0].value)?;
                    let idx = self.import("ecma:set", "has");
                    self.emit_host_call(idx, 2);
                    return Ok(true);
                }
                ("Remove", 1) => {
                    self.compile_expr(object)?;
                    self.compile_collection_key(object, &args[0].value)?;
                    let idx = self.import("ecma:set", "delete");
                    self.emit_host_call(idx, 2);
                    return Ok(true);
                }
                _ => {}
            }
        }

        Ok(false)
    }

    pub(crate) fn resolve_reflection_binding_expr(
        &self,
        expr: &Expression,
    ) -> Option<ReflectionBinding> {
        match &expr.kind {
            ExprKind::Lit(Literal::Str(type_name)) if type_name.starts_with("System.") => {
                Some(ReflectionBinding::Type(type_name.clone()))
            }
            ExprKind::Ident(name) => self.reflection_bindings.get(&self.canon(name)).cloned(),
            ExprKind::Member { object, field, .. } => {
                let receiver = self.resolve_reflection_binding_expr(object)?;
                match (receiver, strip_generic_suffix(field.as_str())) {
                    (ReflectionBinding::Type(type_name), "BaseType") => self
                        .reflection_base_type_name(&type_name)
                        .map(ReflectionBinding::Type),
                    _ => None,
                }
            }
            ExprKind::Call { callee, args, .. } => {
                let ExprKind::Member { object, field, .. } = &callee.kind else {
                    return None;
                };
                let receiver = self.resolve_reflection_binding_expr(object)?;
                match (receiver, strip_generic_suffix(field.as_str())) {
                    (ReflectionBinding::Type(type_name), "GetMethod") => {
                        let method_name = self.resolve_reflection_string_arg(args.first()?)?;
                        Some(ReflectionBinding::Method {
                            type_name,
                            method_name,
                        })
                    }
                    (ReflectionBinding::Type(type_name), "GetProperty") => {
                        let property_name = self.resolve_reflection_string_arg(args.first()?)?;
                        Some(ReflectionBinding::Property {
                            type_name,
                            property_name,
                        })
                    }
                    (ReflectionBinding::Type(type_name), "GetField") => {
                        let field_name = self.resolve_reflection_string_arg(args.first()?)?;
                        Some(ReflectionBinding::Field {
                            type_name,
                            field_name,
                        })
                    }
                    (ReflectionBinding::Type(type_name), "GetNestedType") => {
                        let nested_name = self.resolve_reflection_string_arg(args.first()?)?;
                        self.reflection_nested_type_name(&type_name, &nested_name)
                            .map(ReflectionBinding::Type)
                    }
                    (ReflectionBinding::Type(type_name), "GetGenericTypeDefinition") => Some(
                        ReflectionBinding::Type(self.reflection_open_generic_type_name(&type_name)),
                    ),
                    (ReflectionBinding::Type(type_name), "GetConstructor") => {
                        let param_types =
                            self.resolve_reflection_type_array_expr(&args.first()?.value)?;
                        self.reflection_constructor_for_types(&type_name, &param_types)
                    }
                    _ => None,
                }
            }
            ExprKind::Index { object, index, .. } => {
                let ExprKind::Call { callee, args, .. } = &object.kind else {
                    return None;
                };
                if !args.is_empty() {
                    return None;
                }
                let ExprKind::Member {
                    object: method_object,
                    field,
                    ..
                } = &callee.kind
                else {
                    return None;
                };
                if strip_generic_suffix(field.as_str()) != "GetParameters" {
                    return None;
                }
                let ReflectionBinding::Method {
                    type_name,
                    method_name,
                } = self.resolve_reflection_binding_expr(method_object)?
                else {
                    return None;
                };
                let ExprKind::Lit(Literal::Int(position)) = &index.kind else {
                    return None;
                };
                Some(ReflectionBinding::Parameter {
                    type_name,
                    method_name,
                    index: (*position).max(0) as usize,
                })
            }
            _ => None,
        }
    }

    fn try_compile_js_iterator_from_generator_take_to_array(
        &mut self,
        callee: &Expression,
        args: &[Argument],
    ) -> Result<bool, String> {
        if !self.is_js_profile() || !args.is_empty() {
            return Ok(false);
        }

        let ExprKind::Member {
            object: take_call,
            field: to_array_field,
            null_safe: false,
        } = &callee.kind
        else {
            return Ok(false);
        };
        if to_array_field != "toArray" {
            return Ok(false);
        }

        if let ExprKind::Call {
            callee: flat_map_callee,
            args: flat_map_args,
            optional: false,
        } = &take_call.kind
        {
            if flat_map_args.len() == 1 && !flat_map_args[0].spread {
                if let ExprKind::Member {
                    object: from_call,
                    field: flat_map_field,
                    null_safe: false,
                } = &flat_map_callee.kind
                {
                    let mapper_is_generator = matches!(
                        &flat_map_args[0].value.kind,
                        ExprKind::FunctionExpr(stmt)
                            if matches!(
                                &stmt.kind,
                                StmtKind::FunctionDecl {
                                    is_generator: true,
                                    ..
                                }
                            )
                    );
                    if flat_map_field == "flatMap" && mapper_is_generator {
                        if let ExprKind::Call {
                            callee: from_callee,
                            args: from_args,
                            optional: false,
                        } = &from_call.kind
                        {
                            if from_args.len() == 1
                                && !from_args[0].spread
                                && matches!(&from_args[0].value.kind, ExprKind::Array(_))
                            {
                                if let ExprKind::Member {
                                    object: iterator_obj,
                                    field: from_field,
                                    null_safe: false,
                                } = &from_callee.kind
                                {
                                    if from_field == "from"
                                        && matches!(&iterator_obj.kind, ExprKind::Ident(name) if name == "Iterator")
                                    {
                                        self.compile_expr(&from_args[0].value)?;
                                        self.compile_expr(&flat_map_args[0].value)?;
                                        let line = self.line;
                                        crate::emitter::generators::emit_flat_map_generator_mapper_into_array(
                                            &mut self.chunks,
                                            self.current,
                                            line,
                                        );
                                        return Ok(true);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        if let ExprKind::Call {
            callee: from_callee,
            args: from_args,
            optional: false,
        } = &take_call.kind
        {
            if from_args.len() == 1 && !from_args[0].spread {
                if let ExprKind::Member {
                    object: iterator_obj,
                    field: from_field,
                    null_safe: false,
                } = &from_callee.kind
                {
                    if from_field == "from"
                        && matches!(&iterator_obj.kind, ExprKind::Ident(name) if name == "Iterator")
                    {
                        let source = &from_args[0].value;
                        if self.is_direct_generator_call(source) {
                            self.compile_expr(source)?;
                            let line = self.line;
                            crate::emitter::generators::emit_drain_into_array(
                                &mut self.chunks,
                                self.current,
                                line,
                            );
                            return Ok(true);
                        }
                    }
                }
            }
        }

        let ExprKind::Call {
            callee: take_callee,
            args: take_args,
            optional: false,
        } = &take_call.kind
        else {
            return Ok(false);
        };
        if take_args.len() != 1 || take_args[0].spread {
            return Ok(false);
        }

        let ExprKind::Member {
            object: from_call,
            field: take_field,
            null_safe: false,
        } = &take_callee.kind
        else {
            return Ok(false);
        };
        if take_field != "take" {
            return Ok(false);
        }

        let ExprKind::Call {
            callee: from_callee,
            args: from_args,
            optional: false,
        } = &from_call.kind
        else {
            return Ok(false);
        };
        if from_args.len() != 1 || from_args[0].spread {
            return Ok(false);
        }

        let ExprKind::Member {
            object: iterator_obj,
            field: from_field,
            null_safe: false,
        } = &from_callee.kind
        else {
            return Ok(false);
        };
        if from_field != "from" {
            return Ok(false);
        }
        if !matches!(&iterator_obj.kind, ExprKind::Ident(name) if name == "Iterator") {
            return Ok(false);
        }

        let source = &from_args[0].value;
        if !self.is_direct_generator_call(source) {
            return Ok(false);
        }

        self.compile_expr(source)?;
        self.compile_expr(&take_args[0].value)?;
        let line = self.line;
        crate::emitter::generators::emit_take_into_array(&mut self.chunks, self.current, line);
        Ok(true)
    }

    fn resolve_reflection_string_arg(&self, arg: &Argument) -> Option<String> {
        match &arg.value.kind {
            ExprKind::Lit(Literal::Str(value)) => Some(value.clone()),
            ExprKind::Ident(name) => {
                self.reflection_bindings
                    .get(&self.canon(name))
                    .and_then(|binding| {
                        if let ReflectionBinding::Type(type_name) = binding {
                            Some(type_name.clone())
                        } else {
                            None
                        }
                    })
            }
            _ => None,
        }
    }

    fn resolve_reflection_type_arg(&self, expr: &Expression) -> Option<String> {
        match self.resolve_reflection_binding_expr(expr)? {
            ReflectionBinding::Type(type_name) => Some(type_name),
            _ => None,
        }
    }

    fn resolve_reflection_type_array_expr(&self, expr: &Expression) -> Option<Vec<String>> {
        match &expr.kind {
            ExprKind::Array(items) => items
                .iter()
                .map(|item| self.resolve_reflection_type_arg(&item.value))
                .collect(),
            ExprKind::Lit(Literal::Null) => Some(Vec::new()),
            _ => None,
        }
    }

    fn resolve_reflection_invoke_args(&self, expr: &Expression) -> Option<Vec<Argument>> {
        match &expr.kind {
            ExprKind::Lit(Literal::Null) => Some(Vec::new()),
            ExprKind::Array(items) => Some(
                items
                    .iter()
                    .map(|item| Argument::positional(item.value.clone()))
                    .collect(),
            ),
            _ => None,
        }
    }

    fn resolve_reflection_string_member_expr(&self, expr: &Expression) -> Option<String> {
        let ExprKind::Member { object, field, .. } = &expr.kind else {
            return None;
        };
        match (
            self.resolve_reflection_binding_expr(object)?,
            strip_generic_suffix(field.as_str()),
        ) {
            (ReflectionBinding::Type(type_name), "Name") => {
                Some(self.reflection_type_short_name(&type_name))
            }
            (ReflectionBinding::Type(type_name), "FullName") => {
                Some(self.reflection_type_full_name(&type_name))
            }
            _ => None,
        }
    }

    fn reflection_class_expr(&self, type_name: &str) -> Expression {
        let trimmed = type_name.trim().trim_end_matches('?').trim();
        let without_system = trimmed.strip_prefix("System.").unwrap_or(trimmed);
        let mut parts = without_system.split('.').filter(|part| !part.is_empty());
        let first = parts.next().unwrap_or(without_system);
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

    pub(crate) fn compile_reflection_type_value(&mut self, type_name: &str) -> Result<(), String> {
        let short_name = self.reflection_type_short_name(type_name);
        let full_name = self.reflection_type_full_name(type_name);
        let is_enum = self.reflection_is_enum_type(type_name);
        let is_value_type = self.reflection_is_value_type(type_name);
        self.compile_expr(&Expression::new(ExprKind::Object(vec![
            ObjectProperty::KeyValue {
                key: Expression::string("Name"),
                value: Expression::string(&short_name),
            },
            ObjectProperty::KeyValue {
                key: Expression::string("FullName"),
                value: Expression::string(&full_name),
            },
            ObjectProperty::KeyValue {
                key: Expression::string("IsEnum"),
                value: Expression::bool(is_enum),
            },
            ObjectProperty::KeyValue {
                key: Expression::string("IsValueType"),
                value: Expression::bool(is_value_type),
            },
        ])))
    }

    fn compile_reflection_type_array(&mut self, type_names: &[String]) -> Result<(), String> {
        let line = self.line;
        common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
        for type_name in type_names {
            inst!(self, core_wasm::dup);
            self.compile_reflection_type_value(type_name)?;
            common::collections::emit_push(&mut self.chunks, self.current, line);
            self.emit(Op::DROP);
        }
        Ok(())
    }

    fn reflection_attributes_for_binding(
        &self,
        binding: &ReflectionBinding,
        attribute_type: Option<&str>,
        inherit: bool,
    ) -> Vec<Expression> {
        match binding {
            ReflectionBinding::Type(type_name) => {
                self.reflection_attributes_for_type(type_name, attribute_type, inherit)
            }
            ReflectionBinding::Constructor { .. } => Vec::new(),
            ReflectionBinding::Method {
                type_name,
                method_name,
            } => self
                .reflection_types
                .get(type_name)
                .and_then(|meta| meta.methods.get(method_name))
                .map(|meta| self.filter_reflection_attributes(&meta.decorators, attribute_type))
                .unwrap_or_default(),
            ReflectionBinding::Property {
                type_name,
                property_name,
            } => self
                .reflection_types
                .get(type_name)
                .and_then(|meta| meta.properties.get(property_name))
                .map(|meta| self.filter_reflection_attributes(&meta.decorators, attribute_type))
                .unwrap_or_default(),
            ReflectionBinding::Field {
                type_name,
                field_name,
            } => self
                .reflection_types
                .get(type_name)
                .and_then(|meta| meta.fields.get(field_name))
                .map(|meta| self.filter_reflection_attributes(&meta.decorators, attribute_type))
                .unwrap_or_default(),
            ReflectionBinding::Parameter {
                type_name,
                method_name,
                index,
            } => self
                .reflection_types
                .get(type_name)
                .and_then(|meta| meta.methods.get(method_name))
                .and_then(|meta| meta.params.get(*index))
                .map(|meta| self.filter_reflection_attributes(&meta.decorators, attribute_type))
                .unwrap_or_default(),
        }
    }

    fn reflection_attributes_for_type(
        &self,
        type_name: &str,
        attribute_type: Option<&str>,
        inherit: bool,
    ) -> Vec<Expression> {
        let mut attrs = Vec::new();
        let mut current = Some(type_name.to_string());

        while let Some(current_type) = current {
            let Some(meta) = self.reflection_types.get(&current_type) else {
                break;
            };
            let matching = self.filter_reflection_attributes(&meta.decorators, attribute_type);
            if !matching.is_empty() {
                if let Some(attribute_type) = attribute_type {
                    let usage = self
                        .attribute_usage
                        .get(attribute_type)
                        .copied()
                        .unwrap_or_default();
                    if usage.allow_multiple {
                        attrs.extend(matching);
                    } else {
                        attrs.push(matching[0].clone());
                        break;
                    }
                } else {
                    attrs.extend(matching);
                }
            }

            if !inherit {
                break;
            }
            let should_inherit = attribute_type
                .and_then(|name| self.attribute_usage.get(name))
                .copied()
                .unwrap_or_default()
                .inherited;
            if !should_inherit {
                break;
            }
            current = meta.parents.first().cloned();
        }

        attrs
    }

    fn filter_reflection_attributes(
        &self,
        decorators: &[Expression],
        attribute_type: Option<&str>,
    ) -> Vec<Expression> {
        decorators
            .iter()
            .filter(|decorator| {
                attribute_type.is_none_or(|wanted| {
                    self.reflection_attribute_type_name(decorator)
                        .is_some_and(|actual| actual.eq_ignore_ascii_case(wanted))
                })
            })
            .cloned()
            .collect()
    }

    fn compile_reflection_attribute_instance(&mut self, attr: &Expression) -> Result<(), String> {
        let ExprKind::New { class, args } = &attr.kind else {
            return self.compile_expr(attr);
        };

        let positional_args: Vec<Argument> = args
            .iter()
            .filter(|arg| arg.name.is_none())
            .cloned()
            .collect();
        let named_args: Vec<&Argument> = args.iter().filter(|arg| arg.name.is_some()).collect();
        if named_args.is_empty() {
            return self.compile_expr(attr);
        }

        self.compile_expr(&Expression::new(ExprKind::New {
            class: class.clone(),
            args: positional_args,
        }))?;
        let slot = self.define_local("__reflection_attr");
        self.emit_u16(Op::LOCAL_SET, slot);

        for arg in named_args {
            self.compile_expr(&arg.value)?;
            self.compile_assign_target(&Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident("__reflection_attr")),
                field: arg.name.clone().unwrap_or_default(),
                null_safe: false,
            }))?;
        }

        self.emit_u16(Op::LOCAL_GET, slot);
        Ok(())
    }

    fn compile_reflection_attribute_array(&mut self, attrs: &[Expression]) -> Result<(), String> {
        let line = self.line;
        common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
        for attr in attrs {
            inst!(self, core_wasm::dup);
            self.compile_reflection_attribute_instance(attr)?;
            common::collections::emit_push(&mut self.chunks, self.current, line);
            self.emit(Op::DROP);
        }
        Ok(())
    }

    fn compile_reflection_binding_value(
        &mut self,
        binding: &ReflectionBinding,
    ) -> Result<(), String> {
        match binding {
            ReflectionBinding::Type(type_name) => {
                self.compile_expr(&Expression::new(ExprKind::Object(vec![
                    ObjectProperty::KeyValue {
                        key: Expression::string("Name"),
                        value: Expression::string(type_name),
                    },
                ])))?;
            }
            ReflectionBinding::Constructor { type_name, .. } => {
                self.compile_expr(&Expression::new(ExprKind::Object(vec![
                    ObjectProperty::KeyValue {
                        key: Expression::string("Name"),
                        value: Expression::string(type_name),
                    },
                ])))?;
            }
            ReflectionBinding::Method { method_name, .. } => {
                self.compile_expr(&Expression::new(ExprKind::Object(vec![
                    ObjectProperty::KeyValue {
                        key: Expression::string("Name"),
                        value: Expression::string(method_name),
                    },
                ])))?;
            }
            ReflectionBinding::Property { property_name, .. } => {
                self.compile_expr(&Expression::new(ExprKind::Object(vec![
                    ObjectProperty::KeyValue {
                        key: Expression::string("Name"),
                        value: Expression::string(property_name),
                    },
                ])))?;
            }

            ReflectionBinding::Field { field_name, .. } => {
                self.compile_expr(&Expression::new(ExprKind::Object(vec![
                    ObjectProperty::KeyValue {
                        key: Expression::string("Name"),
                        value: Expression::string(field_name),
                    },
                ])))?;
            }
            ReflectionBinding::Parameter {
                type_name,
                method_name,
                index,
            } => {
                let name = self
                    .reflection_types
                    .get(type_name)
                    .and_then(|meta| meta.methods.get(method_name))
                    .and_then(|meta| meta.params.get(*index))
                    .map(|param| param.name.clone())
                    .unwrap_or_default();
                self.compile_expr(&Expression::new(ExprKind::Object(vec![
                    ObjectProperty::KeyValue {
                        key: Expression::string("Name"),
                        value: Expression::string(&name),
                    },
                    ObjectProperty::KeyValue {
                        key: Expression::string("Position"),
                        value: Expression::int(*index as i64),
                    },
                ])))?;
            }
        }
        Ok(())
    }

    fn try_compile_dotnet_attribute_reflection_call(
        &mut self,
        callee: &Expression,
        args: &[Argument],
    ) -> Result<bool, String> {
        let ExprKind::Member { object, field, .. } = &callee.kind else {
            return Ok(false);
        };
        let field_name = strip_generic_suffix(field);
        let receiver_type = terminal_type_name(object).unwrap_or_default();

        if (receiver_type.eq_ignore_ascii_case("Activator")
            || receiver_type.eq_ignore_ascii_case("System.Activator"))
            && field_name == "CreateInstance"
            && !args.is_empty()
        {
            let Some(type_name) = self.resolve_reflection_type_arg(&args[0].value) else {
                return Ok(false);
            };
            self.compile_expr(&Expression::new(ExprKind::New {
                class: Box::new(self.reflection_class_expr(&type_name)),
                args: Vec::new(),
            }))?;
            return Ok(true);
        }

        if (receiver_type.eq_ignore_ascii_case("Attribute")
            || receiver_type.eq_ignore_ascii_case("System.Attribute"))
            && field_name == "GetCustomAttribute"
            && args.len() >= 2
        {
            let Some(provider) = self.resolve_reflection_binding_expr(&args[0].value) else {
                return Ok(false);
            };
            let Some(attribute_type) = self.resolve_reflection_type_arg(&args[1].value) else {
                return Ok(false);
            };
            let attrs =
                self.reflection_attributes_for_binding(&provider, Some(&attribute_type), true);
            if let Some(attr) = attrs.first() {
                self.compile_reflection_attribute_instance(attr)?;
            } else {
                self.emit(Op::NULL);
            }
            return Ok(true);
        }

        if (receiver_type.eq_ignore_ascii_case("Attribute")
            || receiver_type.eq_ignore_ascii_case("System.Attribute"))
            && field_name == "IsDefined"
            && args.len() >= 2
        {
            let Some(provider) = self.resolve_reflection_binding_expr(&args[0].value) else {
                return Ok(false);
            };
            let Some(attribute_type) = self.resolve_reflection_type_arg(&args[1].value) else {
                return Ok(false);
            };
            let attrs =
                self.reflection_attributes_for_binding(&provider, Some(&attribute_type), true);
            inst!(self, core_wasm::bool_const, !attrs.is_empty());
            return Ok(true);
        }

        let Some(provider) = self.resolve_reflection_binding_expr(object) else {
            return Ok(false);
        };
        match field_name {
            "GetMethod" if args.len() >= 1 => {
                let Some(binding) =
                    self.resolve_reflection_binding_expr(&Expression::new(ExprKind::Call {
                        callee: Box::new(callee.clone()),
                        args: args.to_vec(),
                        optional: false,
                    }))
                else {
                    return Ok(false);
                };
                self.compile_reflection_binding_value(&binding)?;
                Ok(true)
            }
            "GetProperty" if args.len() >= 1 => {
                let Some(binding) =
                    self.resolve_reflection_binding_expr(&Expression::new(ExprKind::Call {
                        callee: Box::new(callee.clone()),
                        args: args.to_vec(),
                        optional: false,
                    }))
                else {
                    return Ok(false);
                };
                self.compile_reflection_binding_value(&binding)?;
                Ok(true)
            }
            "GetField" if args.len() >= 1 => {
                let Some(binding) =
                    self.resolve_reflection_binding_expr(&Expression::new(ExprKind::Call {
                        callee: Box::new(callee.clone()),
                        args: args.to_vec(),
                        optional: false,
                    }))
                else {
                    return Ok(false);
                };
                self.compile_reflection_binding_value(&binding)?;
                Ok(true)
            }
            "GetConstructor" if args.len() >= 1 => {
                let Some(binding) =
                    self.resolve_reflection_binding_expr(&Expression::new(ExprKind::Call {
                        callee: Box::new(callee.clone()),
                        args: args.to_vec(),
                        optional: false,
                    }))
                else {
                    return Ok(false);
                };
                self.compile_reflection_binding_value(&binding)?;
                Ok(true)
            }
            "GetNestedType" if args.len() >= 1 => {
                let Some(binding) =
                    self.resolve_reflection_binding_expr(&Expression::new(ExprKind::Call {
                        callee: Box::new(callee.clone()),
                        args: args.to_vec(),
                        optional: false,
                    }))
                else {
                    self.emit(Op::NULL);
                    return Ok(true);
                };
                self.compile_reflection_binding_value(&binding)?;
                Ok(true)
            }
            "GetGenericArguments" if args.is_empty() => {
                let ReflectionBinding::Type(type_name) = provider else {
                    return Ok(false);
                };
                let args = self.reflection_generic_argument_types(&type_name);
                self.compile_reflection_type_array(&args)?;
                Ok(true)
            }
            "GetGenericTypeDefinition" if args.is_empty() => {
                let Some(binding) =
                    self.resolve_reflection_binding_expr(&Expression::new(ExprKind::Call {
                        callee: Box::new(callee.clone()),
                        args: args.to_vec(),
                        optional: false,
                    }))
                else {
                    return Ok(false);
                };
                self.compile_reflection_binding_value(&binding)?;
                Ok(true)
            }
            "GetInterfaces" if args.is_empty() => {
                let ReflectionBinding::Type(type_name) = provider else {
                    return Ok(false);
                };
                let interfaces = self.reflection_interfaces(&type_name);
                self.compile_reflection_type_array(&interfaces)?;
                Ok(true)
            }
            "IsAssignableFrom" if args.len() >= 1 => {
                let ReflectionBinding::Type(type_name) = provider else {
                    return Ok(false);
                };
                let Some(other_type) = self.resolve_reflection_type_arg(&args[0].value) else {
                    return Ok(false);
                };
                let v = self.reflection_is_assignable_from(&type_name, &other_type);
                inst!(self, core_wasm::bool_const, v);
                Ok(true)
            }
            "GetParameters" if args.is_empty() => {
                let ReflectionBinding::Method {
                    type_name,
                    method_name,
                } = provider
                else {
                    return Ok(false);
                };
                let params = self
                    .reflection_types
                    .get(&type_name)
                    .and_then(|meta| meta.methods.get(&method_name))
                    .map(|meta| meta.params.clone())
                    .unwrap_or_default();
                let line = self.line;
                common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
                for (index, param) in params.iter().enumerate() {
                    inst!(self, core_wasm::dup);
                    self.compile_expr(&Expression::new(ExprKind::Object(vec![
                        ObjectProperty::KeyValue {
                            key: Expression::string("Name"),
                            value: Expression::string(&param.name),
                        },
                        ObjectProperty::KeyValue {
                            key: Expression::string("Position"),
                            value: Expression::int(index as i64),
                        },
                    ])))?;
                    common::collections::emit_push(&mut self.chunks, self.current, line);
                    self.emit(Op::DROP);
                }
                Ok(true)
            }
            "GetCustomAttributes" if args.len() >= 2 => {
                let Some(attribute_type) = self.resolve_reflection_type_arg(&args[0].value) else {
                    return Ok(false);
                };
                let inherit = matches!(args[1].value.kind, ExprKind::Lit(Literal::Bool(true)));
                let attrs = self.reflection_attributes_for_binding(
                    &provider,
                    Some(&attribute_type),
                    inherit,
                );
                self.compile_reflection_attribute_array(&attrs)?;
                Ok(true)
            }
            "Invoke" => match provider {
                ReflectionBinding::Constructor { type_name, .. } => {
                    let ctor_args = args
                        .first()
                        .and_then(|arg| self.resolve_reflection_invoke_args(&arg.value))
                        .unwrap_or_default();
                    self.compile_expr(&Expression::new(ExprKind::New {
                        class: Box::new(self.reflection_class_expr(&type_name)),
                        args: ctor_args,
                    }))?;
                    Ok(true)
                }
                ReflectionBinding::Method { method_name, .. } => {
                    let Some(instance_arg) = args.first() else {
                        return Ok(false);
                    };
                    let invoke_args = args
                        .get(1)
                        .and_then(|arg| self.resolve_reflection_invoke_args(&arg.value))
                        .unwrap_or_default();
                    self.compile_expr(&Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Member {
                            object: Box::new(instance_arg.value.clone()),
                            field: method_name,
                            null_safe: false,
                        })),
                        args: invoke_args,
                        optional: false,
                    }))?;
                    Ok(true)
                }
                _ => Ok(false),
            },
            "GetValue" if !args.is_empty() => match provider {
                ReflectionBinding::Property { property_name, .. }
                | ReflectionBinding::Field {
                    field_name: property_name,
                    ..
                } => {
                    self.compile_expr(&Expression::new(ExprKind::Member {
                        object: Box::new(args[0].value.clone()),
                        field: property_name,
                        null_safe: false,
                    }))?;
                    Ok(true)
                }
                _ => Ok(false),
            },
            "SetValue" if args.len() >= 2 => match provider {
                ReflectionBinding::Property { property_name, .. } => {
                    self.compile_expr(&args[1].value)?;
                    self.compile_assign_target(&Expression::new(ExprKind::Member {
                        object: Box::new(args[0].value.clone()),
                        field: property_name,
                        null_safe: false,
                    }))?;
                    self.emit(Op::NULL);
                    Ok(true)
                }
                ReflectionBinding::Field { field_name, .. } => {
                    if let ExprKind::Ident(name) = &args[0].value.kind {
                        let value_slot = self.define_local("__reflection_field_value");
                        self.compile_expr(&args[1].value)?;
                        self.emit_u16(Op::LOCAL_SET, value_slot);

                        self.compile_expr(&args[0].value)?;
                        inst!(self, core_wasm::dup);
                        self.emit_u16(Op::LOCAL_GET, value_slot);
                        let field_idx = self.str_const(&self.canon(&field_name));
                        self.emit_u16(Op::STRUCT_SET, field_idx);
                        self.emit(Op::DROP);
                        self.emit_var_set(name);
                        self.emit(Op::NULL);
                        Ok(true)
                    } else {
                        self.compile_expr(&args[1].value)?;
                        self.compile_assign_target(&Expression::new(ExprKind::Member {
                            object: Box::new(args[0].value.clone()),
                            field: field_name,
                            null_safe: false,
                        }))?;
                        self.emit(Op::NULL);
                        Ok(true)
                    }
                }
                _ => Ok(false),
            },
            _ => Ok(false),
        }
    }

    fn try_compile_dotnet_delegate_call(
        &mut self,
        callee: &Expression,
        args: &[Argument],
    ) -> Result<bool, String> {
        if args.len() != 2 {
            return Ok(false);
        }

        let ExprKind::Member { object, field, .. } = &callee.kind else {
            return Ok(false);
        };

        let receiver_parts = self.flatten_member_chain(object);
        let Some(receiver_leaf) = receiver_parts.last() else {
            return Ok(false);
        };
        if !receiver_leaf.eq_ignore_ascii_case("Delegate") {
            return Ok(false);
        }

        let emit = if field.eq_ignore_ascii_case("Combine") {
            Some("delegates.combine")
        } else if field.eq_ignore_ascii_case("Remove") {
            Some("delegates.remove")
        } else {
            None
        };
        let Some(emit) = emit else {
            return Ok(false);
        };

        for arg in args {
            self.compile_expr(&arg.value)?;
        }
        let line = self.line;
        self.emit_common(emit, 2, line);
        Ok(true)
    }

    fn try_compile_dotnet_formatted_tostring(
        &mut self,
        callee: &Expression,
        args: &[Argument],
    ) -> Result<bool, String> {
        if args.len() != 1 {
            return Ok(false);
        }
        let ExprKind::Member { object, field, .. } = &callee.kind else {
            return Ok(false);
        };
        if !field.eq_ignore_ascii_case("ToString") {
            return Ok(false);
        }

        let format_looks_string = matches!(&args[0].value.kind, ExprKind::Lit(Literal::Str(_)))
            || resolve_receiver_type_hint(self, &args[0].value)
                .as_deref()
                .map(Self::is_string_type_hint)
                .unwrap_or(false);
        if !format_looks_string {
            return Ok(false);
        }

        let helper = self.str_const("__vybe_dotnet_numeric_format");
        self.emit_u16(Op::GLOBAL_GET, helper);
        self.compile_expr(object)?;
        self.compile_expr(&args[0].value)?;
        self.emit_const(Value::F64(0.0));
        self.emit_u8(Op::CALL_REF, 3);
        Ok(true)
    }

    fn try_compile_dotnet_zero_arg_tostring(
        &mut self,
        callee: &Expression,
        args: &[Argument],
    ) -> Result<bool, String> {
        if !matches!(self.profile.name.as_str(), "csharp" | "vb")
            || !self.profile.namespaces.use_dotnet
        {
            return Ok(false);
        }
        if !args.is_empty() {
            return Ok(false);
        }
        let ExprKind::Member { object, field, .. } = &callee.kind else {
            return Ok(false);
        };
        if !field.eq_ignore_ascii_case("ToString") {
            return Ok(false);
        }

        if let Some(class_name) = resolve_receiver_type_hint(self, object) {
            let class_name = Self::normalize_type_hint(&class_name);
            if let Some(target) =
                common::dotnet::surface().lookup_instance_method(&class_name, field, 0)
            {
                self.compile_expr(object)?;
                match target {
                    common::dotnet::InstanceMethodTarget::Host { module, func, .. } => {
                        let idx = self.import(&module, &func);
                        self.emit_host_call(idx, 1);
                    }
                    common::dotnet::InstanceMethodTarget::Common { emit, .. } => {
                        let line = self.line;
                        self.emit_common(&emit, 1, line);
                    }
                }
                return Ok(true);
            }
        }

        self.compile_expr(object)?;
        let obj_slot = self.define_local("__dotnet_tostring_obj");
        self.emit_u16(Op::LOCAL_SET, obj_slot);

        self.emit_u16(Op::LOCAL_GET, obj_slot);
        fn_call!(self, "ecma:value", "typeof", 1);
        let type_slot = self.define_local("__dotnet_tostring_type");
        self.emit_u16(Op::LOCAL_SET, type_slot);

        let result_slot = self.define_local("__dotnet_tostring_result");
        self.emit(Op::NULL);
        self.emit_u16(Op::LOCAL_SET, result_slot);

        let primitive_slot = self.define_local("__dotnet_tostring_primitive");
        self.emit_const(Value::I32(0));
        self.emit_u16(Op::LOCAL_SET, primitive_slot);

        for type_name in ["number", "i32", "i64", "string", "boolean"] {
            self.emit_u16(Op::LOCAL_GET, primitive_slot);
            self.emit(Op::I32_EQZ);
            let line = self.line;
            self.chunk().emit_if(line);
            self.emit_u16(Op::LOCAL_GET, type_slot);
            self.emit_const(Value::String(Arc::from(type_name)));
            {
                let line = self.line;
                crate::emitter::ops::emit_dyn_eq(self.chunk(), line);
            };
            let line = self.line;
            crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
            self.chunk().emit_if(line);
            self.emit_const(Value::I32(1));
            self.emit_u16(Op::LOCAL_SET, primitive_slot);
            self.chunk().emit_end(line);
            self.chunk().emit_end(line);
        }

        self.emit_u16(Op::LOCAL_GET, primitive_slot);
        let line = self.line;
        self.chunk().emit_if(line);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        common::strings::emit_to_string(self.chunk(), line);
        self.emit_u16(Op::LOCAL_SET, result_slot);
        self.chunk().emit_else(line);

        let canon_key = self.str_const(&self.canon(field));
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_u16(Op::STRUCT_GET, canon_key);
        let fn_slot = self.define_local("__dotnet_tostring_fn");
        self.emit_u16(Op::LOCAL_SET, fn_slot);

        self.emit_u16(Op::LOCAL_GET, fn_slot);
        fn_call!(self, "wasm:js-undefined", "test", 1);
        let line = self.line;
        self.chunk().emit_if(line);
        if field.as_str() != self.canon(field) {
            let exact_key = self.str_const(field);
            self.emit_u16(Op::LOCAL_GET, obj_slot);
            self.emit_u16(Op::STRUCT_GET, exact_key);
            self.emit_u16(Op::LOCAL_SET, fn_slot);
        }
        self.chunk().emit_end(line);

        self.emit_u16(Op::LOCAL_GET, fn_slot);
        fn_call!(self, "wasm:js-undefined", "test", 1);
        let line = self.line;
        self.chunk().emit_if(line);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        let type_key = self.str_const("__type");
        self.emit_u16(Op::STRUCT_GET, type_key);
        self.emit_const(Value::String(Arc::from("Guid")));
        {
            let line = self.line;
            crate::emitter::ops::emit_dyn_eq(self.chunk(), line);
        };
        let line = self.line;
        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if_value(line);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        let value_key = self.str_const("__value");
        self.emit_u16(Op::STRUCT_GET, value_key);
        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        common::strings::emit_to_string(self.chunk(), line);
        self.chunk().emit_end(line);
        self.emit_u16(Op::LOCAL_SET, result_slot);

        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, fn_slot);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_u8(Op::CALL_REF, 1);
        self.emit_u16(Op::LOCAL_SET, result_slot);
        self.chunk().emit_end(line);
        self.chunk().emit_end(line);
        self.emit_u16(Op::LOCAL_GET, result_slot);
        Ok(true)
    }

    fn canonical_enum_type_from_runtime_type(&self, expr: &Expression) -> Option<String> {
        let ExprKind::Lit(Literal::Str(type_name)) = &expr.kind else {
            return None;
        };
        let short = type_name.rsplit('.').next().unwrap_or(type_name).trim();
        self.resolve_known_enum_type(short)
    }

    pub(super) fn canonical_enum_type_from_expr(&self, expr: &Expression) -> Option<String> {
        match &expr.kind {
            ExprKind::Ident(name) => self
                .lookup_var_type_hint(name)
                .and_then(|hint| self.resolve_known_enum_type(hint))
                .or_else(|| self.resolve_known_enum_type(name)),
            ExprKind::Member { object, .. } => {
                let enum_type = terminal_type_name(object)?;
                self.resolve_known_enum_type(strip_generic_suffix(&enum_type))
            }
            _ => resolve_receiver_type_hint(self, expr)
                .and_then(|hint| self.resolve_known_enum_type(strip_generic_suffix(&hint))),
        }
    }

    pub(super) fn console_enum_type_from_expr(&self, expr: &Expression) -> Option<String> {
        match &expr.kind {
            ExprKind::Ident(_) => self.canonical_enum_type_from_expr(expr),
            ExprKind::Member { object, .. } if !matches!(&object.kind, ExprKind::Ident(_)) => {
                self.canonical_enum_type_from_expr(expr)
            }
            _ => None,
        }
    }

    pub(super) fn resolve_known_enum_type(&self, name: &str) -> Option<String> {
        let canon = self.canon(name);
        if self.enum_value_names.contains_key(&canon) {
            return Some(canon);
        }
        self.enum_value_names
            .keys()
            .find(|known| known.eq_ignore_ascii_case(name) || known.eq_ignore_ascii_case(&canon))
            .cloned()
    }

    pub(super) fn enum_member_ordinal(&self, enum_type: &str, member_name: &str) -> Option<i64> {
        let enum_type = self.resolve_known_enum_type(enum_type)?;
        self.enum_value_names
            .get(&enum_type)?
            .iter()
            .find(|(_, name)| name.eq_ignore_ascii_case(member_name))
            .map(|(value, _)| *value)
    }

    fn enum_entries_sorted(&self, enum_type: &str) -> Option<Vec<(i64, String)>> {
        let mut entries: Vec<(i64, String)> = self
            .enum_value_names
            .get(enum_type)?
            .iter()
            .map(|(value, name)| (*value, name.clone()))
            .collect();
        entries.sort_by_key(|(value, _)| *value);
        Some(entries)
    }

    fn compile_string_array(&mut self, values: &[String]) -> Result<(), String> {
        let expr = Expression::new(ExprKind::Array(
            values
                .iter()
                .map(|value| ArrayElement {
                    key: None,
                    value: Expression::string(value),
                    spread: false,
                    by_ref: false,
                })
                .collect(),
        ));
        self.compile_expr(&expr)
    }

    fn emit_enum_name_lookup(
        &mut self,
        enum_type: &str,
        value_expr: &Expression,
        ignore_case: bool,
    ) -> Result<(), String> {
        let Some(entries) = self.enum_entries_sorted(enum_type) else {
            self.emit(Op::NULL);
            return Ok(());
        };

        let line = self.line;
        let to_str_idx = self.import("ecma:string", "String");
        let lower_idx = if ignore_case {
            Some(self.import("ecma:string", "toLowerCase"))
        } else {
            None
        };

        self.compile_expr(value_expr)?;
        self.emit_host_call(to_str_idx, 1);
        if let Some(lower_idx) = lower_idx {
            self.emit_host_call(lower_idx, 1);
        }
        let input_slot = self.define_local("__enum_name_input");
        self.emit_u16(Op::LOCAL_SET, input_slot);

        let result_slot = self.define_local("__enum_name_result");
        let matched_slot = self.define_local("__enum_name_matched");
        self.emit(Op::NULL);
        self.emit_u16(Op::LOCAL_SET, result_slot);
        self.emit_const(Value::I32(0));
        self.emit_u16(Op::LOCAL_SET, matched_slot);
        for (_, name) in entries {
            self.emit_u16(Op::LOCAL_GET, matched_slot);
            self.emit(Op::I32_EQZ);
            let line = self.line;
            self.chunk().emit_if(line);
            self.emit_u16(Op::LOCAL_GET, input_slot);
            let candidate = if ignore_case {
                name.to_ascii_lowercase()
            } else {
                name.clone()
            };
            self.emit_const(Value::String(Arc::from(candidate.as_str())));
            {
                let line = self.line;
                crate::emitter::ops::emit_dyn_eq(self.chunk(), line);
            };
            let line = self.line;
            crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
            self.chunk().emit_if(line);
            self.emit_const(Value::String(Arc::from(name.as_str())));
            self.emit_u16(Op::LOCAL_SET, result_slot);
            self.emit_const(Value::I32(1));
            self.emit_u16(Op::LOCAL_SET, matched_slot);
            self.chunk().emit_end(line);
            self.chunk().emit_end(line);
        }

        self.emit_u16(Op::LOCAL_GET, result_slot);
        let _ = line;
        Ok(())
    }

    pub(super) fn emit_enum_value_to_string(
        &mut self,
        enum_type: &str,
        value_expr: &Expression,
    ) -> Result<(), String> {
        let Some(entries) = self.enum_entries_sorted(enum_type) else {
            self.compile_expr(value_expr)?;
            let to_str_idx = self.import("ecma:string", "String");
            self.emit_host_call(to_str_idx, 1);
            return Ok(());
        };

        let value_slot = self.define_local("__enum_tostring_value");
        self.compile_expr(value_expr)?;
        self.emit_u16(Op::LOCAL_SET, value_slot);

        let result_slot = self.define_local("__enum_tostring_result");
        let matched_slot = self.define_local("__enum_tostring_matched");
        self.emit(Op::NULL);
        self.emit_u16(Op::LOCAL_SET, result_slot);
        self.emit_const(Value::I32(0));
        self.emit_u16(Op::LOCAL_SET, matched_slot);
        for (value, name) in &entries {
            self.emit_u16(Op::LOCAL_GET, matched_slot);
            self.emit(Op::I32_EQZ);
            let line = self.line;
            self.chunk().emit_if(line);
            self.emit_u16(Op::LOCAL_GET, value_slot);
            self.emit_const(Value::F64(*value as f64));
            {
                let line = self.line;
                crate::emitter::ops::emit_dyn_eq(self.chunk(), line);
            };
            let line = self.line;
            crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
            self.chunk().emit_if(line);
            self.emit_const(Value::String(Arc::from(name.as_str())));
            self.emit_u16(Op::LOCAL_SET, result_slot);
            self.emit_const(Value::I32(1));
            self.emit_u16(Op::LOCAL_SET, matched_slot);
            self.chunk().emit_end(line);
            self.chunk().emit_end(line);
        }

        if self.enum_flags.contains(enum_type) {
            self.emit_u16(Op::LOCAL_GET, matched_slot);
            self.emit(Op::I32_EQZ);
            let line = self.line;
            self.chunk().emit_if(line);
            self.emit_const(Value::String(Arc::from("")));
            self.emit_u16(Op::LOCAL_SET, result_slot);
            self.emit_const(Value::I32(0));
            self.emit_u16(Op::LOCAL_SET, matched_slot);

            for (value, name) in &entries {
                if *value <= 0 || (value & (value - 1)) != 0 {
                    continue;
                }
                self.emit_u16(Op::LOCAL_GET, value_slot);
                self.emit_const(Value::F64(*value as f64));
                self.emit(Op::I32_AND);
                self.emit_const(Value::F64(*value as f64));
                {
                    let line = self.line;
                    crate::emitter::ops::emit_dyn_eq(self.chunk(), line);
                };
                let line = self.line;
                crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                self.chunk().emit_if(line);

                self.emit_u16(Op::LOCAL_GET, matched_slot);
                crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
                self.chunk().emit_if_value(line);
                self.emit_u16(Op::LOCAL_GET, result_slot);
                self.emit_const(Value::String(Arc::from(", ")));
                {
                    let line = self.line;
                    crate::emitter::ops::emit_dyn_add(self.chunk(), line);
                };
                self.emit_const(Value::String(Arc::from(name.as_str())));
                {
                    let line = self.line;
                    crate::emitter::ops::emit_dyn_add(self.chunk(), line);
                };
                self.chunk().emit_else(line);
                self.emit_const(Value::String(Arc::from(name.as_str())));
                self.chunk().emit_end(line);

                self.emit_u16(Op::LOCAL_SET, result_slot);
                self.emit_const(Value::I32(1));
                self.emit_u16(Op::LOCAL_SET, matched_slot);
                self.chunk().emit_end(line);
            }

            self.emit_u16(Op::LOCAL_GET, matched_slot);
            let line = self.line;
            crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
            self.chunk().emit_if(line);
            self.emit_u16(Op::LOCAL_GET, result_slot);
            self.emit_u16(Op::LOCAL_SET, result_slot);
            self.chunk().emit_end(line);
            self.chunk().emit_end(line);
        }

        self.emit_u16(Op::LOCAL_GET, matched_slot);
        let line = self.line;
        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if_value(line);
        self.emit_u16(Op::LOCAL_GET, result_slot);
        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        let to_str_idx = self.import("ecma:string", "String");
        self.emit_host_call(to_str_idx, 1);
        self.chunk().emit_end(line);
        Ok(())
    }

    pub(super) fn emit_dotnet_console_arg(&mut self, expr: &Expression) -> Result<(), String> {
        if let Some(enum_type) = self.console_enum_type_from_expr(expr) {
            self.emit_enum_value_to_string(&enum_type, expr)?;
            return Ok(());
        }

        if self.profile.name != "csharp" {
            self.compile_expr(expr)?;
            return Ok(());
        }

        self.compile_expr(expr)?;
        let value_slot = self.define_local("__dotnet_console_value");
        self.emit_u16(Op::LOCAL_SET, value_slot);

        self.emit_u16(Op::LOCAL_GET, value_slot);
        fn_call!(self, "ecma:value", "typeof", 1);
        self.emit_const(Value::String(Arc::from("number")));
        {
            let line = self.line;
            crate::emitter::ops::emit_dyn_eq(self.chunk(), line);
        };
        let line = self.line;
        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if_value(line);

        let helper = self.str_const("__vybe_dotnet_numeric_format");
        self.emit_u16(Op::GLOBAL_GET, helper);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit_const(Value::String(Arc::from("F12")));
        self.emit_const(Value::F64(0.0));
        self.emit_u8(Op::CALL_REF, 3);
        let parse_float = self.import("ecma:number", "parseFloat");
        self.emit_host_call(parse_float, 1);
        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.chunk().emit_end(line);
        Ok(())
    }

    fn emit_enum_has_flag(
        &mut self,
        value_expr: &Expression,
        flag_expr: &Expression,
    ) -> Result<(), String> {
        let flag_slot = self.define_local("__enum_flag_value");
        let value_slot = self.define_local("__enum_flag_source");
        self.compile_expr(flag_expr)?;
        self.emit_u16(Op::LOCAL_SET, flag_slot);
        self.compile_expr(value_expr)?;
        self.emit_u16(Op::LOCAL_SET, value_slot);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit_u16(Op::LOCAL_GET, flag_slot);
        self.emit(Op::I32_AND);
        self.emit_u16(Op::LOCAL_GET, flag_slot);
        {
            let line = self.line;
            crate::emitter::ops::emit_dyn_eq(self.chunk(), line);
        };
        Ok(())
    }

    fn try_compile_dotnet_enum_call(
        &mut self,
        callee: &Expression,
        args: &[Argument],
    ) -> Result<bool, String> {
        let mut static_enum_call = false;
        let (field, instance_object) = match &callee.kind {
            ExprKind::Member { object, field, .. } => {
                if terminal_type_name(object)
                    .is_some_and(|type_name| type_name.eq_ignore_ascii_case("Enum"))
                {
                    static_enum_call = true;
                    (field.as_str(), None)
                } else {
                    (field.as_str(), Some(object.as_ref()))
                }
            }
            ExprKind::Ident(name) => {
                let Some((receiver, field)) = name.rsplit_once('.') else {
                    return Ok(false);
                };
                if receiver
                    .rsplit('.')
                    .next()
                    .is_some_and(|type_name| type_name.eq_ignore_ascii_case("Enum"))
                {
                    static_enum_call = true;
                    (field, None)
                } else {
                    return Ok(false);
                }
            }
            _ => return Ok(false),
        };
        let field_name = strip_generic_suffix(field);

        if static_enum_call {
            match field_name {
                "GetNames" if args.len() == 1 => {
                    let Some(enum_type) =
                        self.canonical_enum_type_from_runtime_type(&args[0].value)
                    else {
                        return Ok(false);
                    };
                    let Some(entries) = self.enum_entries_sorted(&enum_type) else {
                        return Ok(false);
                    };
                    let names: Vec<String> = entries.into_iter().map(|(_, name)| name).collect();
                    self.compile_string_array(&names)?;
                    return Ok(true);
                }
                "GetValues" if args.len() == 1 => {
                    let Some(enum_type) =
                        self.canonical_enum_type_from_runtime_type(&args[0].value)
                    else {
                        return Ok(false);
                    };
                    let Some(entries) = self.enum_entries_sorted(&enum_type) else {
                        return Ok(false);
                    };
                    let names: Vec<String> = entries.into_iter().map(|(_, name)| name).collect();
                    self.compile_string_array(&names)?;
                    return Ok(true);
                }
                "Parse" if args.len() >= 2 => {
                    let Some(enum_type) =
                        self.canonical_enum_type_from_runtime_type(&args[0].value)
                    else {
                        return Ok(false);
                    };
                    self.emit_enum_name_lookup(&enum_type, &args[1].value, false)?;
                    return Ok(true);
                }
                "IsDefined" if args.len() >= 2 => {
                    let Some(enum_type) =
                        self.canonical_enum_type_from_runtime_type(&args[0].value)
                    else {
                        return Ok(false);
                    };
                    self.emit_enum_name_lookup(&enum_type, &args[1].value, false)?;
                    self.emit(Op::REF_IS_NULL);
                    {
                        let line = self.line;
                        crate::emitter::ops::emit_dyn_not(self.chunk(), line);
                    };
                    return Ok(true);
                }
                "GetUnderlyingType" if args.len() == 1 => {
                    let expr = Expression::new(ExprKind::Object(vec![
                        ObjectProperty::KeyValue {
                            key: Expression::string("Name"),
                            value: Expression::string("Int32"),
                        },
                        ObjectProperty::KeyValue {
                            key: Expression::string("FullName"),
                            value: Expression::string("System.Int32"),
                        },
                    ]));
                    self.compile_expr(&expr)?;
                    return Ok(true);
                }
                "Format" if args.len() >= 3 => {
                    self.compile_expr(&args[1].value)?;
                    let to_str_idx = self.import("ecma:string", "String");
                    self.emit_host_call(to_str_idx, 1);
                    return Ok(true);
                }
                "TryParse" if matches!(args.len(), 2 | 3 | 4 | 5) => {
                    let visible_args = if args.len() >= 4 {
                        &args[..args.len() - 2]
                    } else {
                        args
                    };
                    let enum_type = extract_generic_type_name(field)
                        .map(|name| self.canon(&name))
                        .filter(|canon| self.enum_value_names.contains_key(canon))
                        .or_else(|| {
                            (args.len() >= 4)
                                .then(|| {
                                    self.canonical_enum_type_from_expr(&args[args.len() - 2].value)
                                })
                                .flatten()
                        });
                    let Some(enum_type) = enum_type else {
                        return Ok(false);
                    };
                    let (value_arg, ignore_case, out_arg) = if visible_args.len() == 3 {
                        (
                            &visible_args[0].value,
                            matches!(
                                visible_args[1].value.kind,
                                ExprKind::Lit(Literal::Bool(true))
                            ),
                            &visible_args[2].value,
                        )
                    } else {
                        (&visible_args[0].value, false, &visible_args[1].value)
                    };
                    self.emit_enum_name_lookup(&enum_type, value_arg, ignore_case)?;
                    let parsed_slot = self.define_local("__enum_try_parse_value");
                    self.emit_u16(Op::LOCAL_SET, parsed_slot);
                    self.emit_u16(Op::LOCAL_GET, parsed_slot);
                    self.emit(Op::REF_IS_NULL);
                    let line = self.line;
                    self.chunk().emit_if_value(line);
                    self.emit(Op::NULL);
                    self.compile_assign_target(out_arg)?;
                    inst!(self, core_wasm::bool_const, false);
                    self.chunk().emit_else(line);
                    self.emit_u16(Op::LOCAL_GET, parsed_slot);
                    self.compile_assign_target(out_arg)?;
                    inst!(self, core_wasm::bool_const, true);
                    self.chunk().emit_end(line);
                    return Ok(true);
                }
                _ => {}
            }
        }

        let Some(object) = instance_object else {
            return Ok(false);
        };

        let Some(enum_type) = self.canonical_enum_type_from_expr(object) else {
            return Ok(false);
        };

        match field_name {
            "HasFlag" if args.len() == 1 => {
                self.emit_enum_has_flag(object, &args[0].value)?;
                Ok(true)
            }
            "ToString" if args.is_empty() => {
                self.emit_enum_value_to_string(&enum_type, object)?;
                Ok(true)
            }
            _ => Ok(false),
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    // Lambda compilation
    // ════════════════════════════════════════════════════════════════════════

    fn split_explicit_capture(capture: &str) -> (bool, &str) {
        if let Some(name) = capture.strip_prefix('&') {
            (true, name)
        } else {
            (false, capture)
        }
    }

    fn normalize_explicit_capture(&self, capture: &str) -> String {
        let (by_ref, raw_name) = Self::split_explicit_capture(capture);
        let normalized_name = if self.is_php_profile() && !raw_name.starts_with('$') {
            format!("${raw_name}")
        } else {
            raw_name.to_string()
        };

        if by_ref {
            format!("&{normalized_name}")
        } else {
            normalized_name
        }
    }

    pub(super) fn compile_lambda(
        &mut self,
        params: &[Param],
        body: &LambdaBody,
        captures: &[String],
    ) -> Result<(), String> {
        self.compile_lambda_with_flags(params, body, captures, false, false, false)
    }

    pub(super) fn compile_lambda_with_flags(
        &mut self,
        params: &[Param],
        body: &LambdaBody,
        captures: &[String],
        is_async: bool,
        is_generator: bool,
        is_arrow: bool,
    ) -> Result<(), String> {
        let normalized_captures: Vec<String> = captures
            .iter()
            .map(|capture| self.normalize_explicit_capture(capture))
            .collect();

        if normalized_captures
            .iter()
            .any(|capture| !Self::split_explicit_capture(capture).0)
        {
            return self.compile_lambda_with_explicit_captures(
                params,
                body,
                &normalized_captures,
                is_async,
                is_generator,
            );
        }

        self.compile_lambda_direct(params, body, is_async, is_generator, is_arrow)
    }

    fn compile_lambda_with_explicit_captures(
        &mut self,
        params: &[Param],
        body: &LambdaBody,
        captures: &[String],
        is_async: bool,
        is_generator: bool,
    ) -> Result<(), String> {
        let capture_bindings: Vec<(String, Option<String>)> = captures
            .iter()
            .filter_map(|capture| {
                let (by_ref, capture_name) = Self::split_explicit_capture(capture);
                if by_ref {
                    None
                } else {
                    Some((
                        capture_name.to_string(),
                        self.lookup_var_type_hint(capture_name).map(str::to_string),
                    ))
                }
            })
            .collect();

        let factory_idx = self.chunks.len();
        let factory = common::functions::create_function_chunk(
            "<lambda_factory>",
            capture_bindings.len() as u8,
        );
        self.chunks.push(factory);
        self.scopes.push(Scope::new_function());
        let saved = self.current;
        self.current = factory_idx;

        for (capture_name, capture_type) in &capture_bindings {
            self.define_local_typed(capture_name, capture_type.clone());
        }

        // Compile the actual lambda body inside the factory. The inner lambda
        // upvalue-captures the factory's locals (the by-value captures, including
        // __js_this). compile_lambda_direct emits REF_FUNC into the factory chunk,
        // leaving the function reference on the factory's operand stack.
        // PHP `use` closures — never arrows.
        self.compile_lambda_direct(params, body, is_async, is_generator, false)?;

        // Emit RETURN so the factory returns the function reference it just built.
        let line = self.line;
        self.chunks[factory_idx].emit_op(Op::RETURN, line);

        // Collect upvalues AFTER body compilation — the body may have referenced
        // outer-scope variables, registering them as factory upvalues.
        let uvs = self.scopes.last().unwrap().upvalues.clone();
        let inner_scope_idx = self.scopes.len() - 1;
        let uv_names: Vec<Option<String>> = (0..uvs.len())
            .map(|i| self.captured_name_for_upvalue(inner_scope_idx, i as u8))
            .collect();
        self.scopes.pop();
        self.current = saved;

        let line = self.line;
        if uvs.is_empty() {
            common::functions::emit_ref_func(&mut self.chunks[self.current], factory_idx, 0, line);
        } else {
            let mut env_slots: Vec<u16> = Vec::new();
            for (i, uv) in uvs.iter().enumerate() {
                if let Some(name) = uv_names[i].clone() {
                    let slot = if uv.is_local {
                        uv.index as u16
                    } else {
                        let parent_env = self.closure_env_slot();
                        let parent_idx = self.closure_env_index(&name);
                        let tmp = self.define_local(&format!("__nested_cap_{}", name));
                        crate::emitter::closures::emit_env_get(
                            self.chunk(),
                            parent_env,
                            parent_idx,
                            line,
                        );
                        self.emit_u16(Op::LOCAL_SET, tmp);
                        tmp
                    };
                    env_slots.push(slot);
                }
            }
            crate::emitter::closures::emit_env_new(self.chunk(), &env_slots, line);
            let env_slot = self.define_local(&format!("__closure_env_factory_{}", factory_idx));
            self.emit_u16(Op::LOCAL_SET, env_slot);
            common::functions::emit_ref_func(&mut self.chunks[self.current], factory_idx, 1, line);
            self.chunks[self.current].emit(1, line);
            self.chunks[self.current].emit(env_slot as u8, line);
        }
        for capture in captures {
            let (by_ref, capture_name) = Self::split_explicit_capture(capture);
            if !by_ref {
                if self.is_js_profile() && capture_name == "__js_this" {
                    self.compile_expr(&Expression::new(ExprKind::This))?;
                } else {
                    self.emit_var_get(capture_name);
                }
            }
        }
        self.emit_u8(Op::CALL_REF, capture_bindings.len() as u8);
        Ok(())
    }

    fn compile_lambda_direct(
        &mut self,
        params: &[Param],
        body: &LambdaBody,
        is_async: bool,
        is_generator: bool,
        is_arrow: bool,
    ) -> Result<(), String> {
        // Walker-lowered generator/async-generator expressions arrive as
        // Lambdas holding the `__gen_fn` contract — they are NOT arrows.
        let is_arrow = is_arrow
            && match body {
                LambdaBody::Block(stmts) => Self::wrapped_generator_kind(stmts).is_none(),
                _ => true,
            };
        let has_rest = params.last().map_or(false, |p| p.is_rest);
        if has_rest {
            self.rest_fixed_arities
                .insert(params.len().saturating_sub(1) as u8);
        }
        // §10.2.11 / §15.3: arrows bind `this` and `new.target` LEXICALLY —
        // captured at CREATION, never read from the ambient call-time
        // globals (member calls set __js_this to the receiver, plain calls
        // null __js_new_target; both are [[Call]]-time bindings arrows must
        // not observe). When the enclosing scope already provides a lexical
        // `this` (method/ctor local, or an outer arrow's capture) the
        // existing upvalue resolution is the capture; otherwise snapshot
        // the current globals into enclosing locals the arrow body's
        // upvalue resolution will find by name.
        if self.is_js_profile() && is_arrow {
            let self_kw = self.profile.self_keyword.clone();
            let scope_idx = self.scopes.len() - 1;
            let this_reachable = self.scope().resolve(&self_kw).is_some()
                || self.scope().resolve("__js_this").is_some()
                || (scope_idx > 0
                    && (self.resolve_upvalue(scope_idx, &self_kw).is_some()
                        || self.resolve_upvalue(scope_idx, "__js_this").is_some()));
            if !this_reachable {
                let slot = self.define_local("__js_this");
                let js_this = self.str_const("__js_this");
                self.emit_u16(Op::GLOBAL_GET, js_this);
                self.emit_u16(Op::LOCAL_SET, slot);
            }
            let nt_reachable = self.scope().resolve("__js_new_target").is_some()
                || (scope_idx > 0 && self.resolve_upvalue(scope_idx, "__js_new_target").is_some());
            if !nt_reachable {
                let slot = self.define_local("__js_new_target");
                let js_nt = self.str_const("__js_new_target");
                self.emit_u16(Op::GLOBAL_GET, js_nt);
                self.emit_u16(Op::LOCAL_SET, slot);
            }
        }
        // Capture parent's shared env info before switching scope
        let parent_shared_env_slot = self.shared_env_slot;
        let parent_shared_env_names = self.shared_env_names.clone();
        let arity = params.len() as u8;
        let ci = self.chunks.len();
        let chunk = common::functions::create_function_chunk("<lambda>", arity);
        self.chunks.push(chunk);
        self.chunks[ci].is_async = is_async;
        self.chunks[ci].is_generator = is_generator;
        self.scopes.push(Scope::new_function());
        let saved = self.current;
        self.current = ci;
        // Runtime TRY_END counts are per-FRAME: a nested chunk must not
        // inherit the enclosing async body's try depth, or its returns pop the
        // CALLER's handlers off the shared runtime handler stack (a lambda
        // compiled inline inside an async fn emitted TRY_END × 2, silently
        // removing the user's enclosing try/catch).
        let saved_async_try_depth = std::mem::take(&mut self.active_async_try_depth);
        let saved_fn = self.current_func_name.replace("<lambda>".into());
        let saved_env_names = std::mem::take(&mut self.closure_env_names);
        let saved_capture_locals = std::mem::take(&mut self.capture_locals);
        let saved_shared_env_slot = self.shared_env_slot.take();
        let saved_shared_env_names = std::mem::take(&mut self.shared_env_names);
        // If parent has a shared env, pre-seed the inner function's
        // closure_env_names so upvalue indices match the shared env layout.
        if !parent_shared_env_names.is_empty() {
            self.closure_env_names = parent_shared_env_names.clone();
        }
        // ECMA-262 §11.2.2: strict mode is inherited by nested functions and
        // additionally turned on by a `"use strict"` directive prologue in
        // this function's own block body. Arrow expression bodies cannot carry
        // a prologue, so they only inherit.
        let saved_strict = self.in_strict;
        let saved_closure_captured = std::mem::take(&mut self.current_closure_captured_locals);
        match body {
            LambdaBody::Block(stmts) => {
                if Self::stmts_have_use_strict_directive(stmts) {
                    self.in_strict = true;
                }
                crate::compiler::collect_closure_captured_idents(
                    stmts,
                    &mut self.current_closure_captured_locals,
                );
            }
            LambdaBody::Expr(expr) => {
                crate::compiler::collect_closure_captured_in_expr(
                    expr,
                    &mut self.current_closure_captured_locals,
                );
            }
        }
        for p in params {
            self.define_local_typed(&p.name, p.type_hint.clone());
            if let Some(ref default) = p.default {
                let slot = self.scope().resolve(&p.name).unwrap();
                self.emit_u16(Op::LOCAL_GET, slot);
                self.emit(Op::REF_IS_NULL);
                let line = self.line;
                self.chunk().emit_if(line);
                self.compile_expr(default)?;
                self.emit_u16(Op::LOCAL_SET, slot);
                self.chunk().emit_end(line);
            }
        }
        // Snapshot __js_this as a local BEFORE shared env creation so inner
        // arrows can capture it via the shared env / upvalue chain.
        if self.is_js_profile() && self.scopes.len() > 1 {
            let parent_has_this = self.scopes.len() > 2
                && self.scopes[self.scopes.len() - 2]
                    .resolve("__js_this")
                    .is_some();
            if !parent_has_this {
                let body_has_this = match body {
                    LambdaBody::Block(stmts) => crate::compiler::body_contains_this(stmts),
                    LambdaBody::Expr(expr) => crate::compiler::expr_contains_this(expr),
                };
                if body_has_this {
                    let this_idx = self.str_const("__js_this");
                    self.emit_u16(Op::GLOBAL_GET, this_idx);
                    let this_local = self.define_local("__js_this");
                    self.emit_u16(Op::LOCAL_SET, this_local);
                    self.current_closure_captured_locals
                        .insert("__js_this".to_string());
                }
            }
        }

        if !self.current_closure_captured_locals.is_empty() {
            let mut captured_names: Vec<String> = self
                .current_closure_captured_locals
                .iter()
                .filter(|name| !self.defined_globals.contains(name.as_str()))
                .cloned()
                .collect();
            captured_names.sort();

            {
                let env_size = captured_names.len() as u16;
                let line = self.line;
                for _ in 0..env_size {
                    self.emit(Op::NULL);
                }
                self.chunks[self.current].emit_op_u16(Op::ARRAY_NEW_FIXED, env_size, line);
                let env_slot = self.define_local("__shared_env");
                self.emit_u16(Op::LOCAL_SET, env_slot);
                self.shared_env_slot = Some(env_slot);
                self.shared_env_names = captured_names.clone();
                let mut local_decls: HashSet<String> =
                    params.iter().map(|p| p.name.clone()).collect();
                local_decls.insert("__js_this".to_string());
                if let LambdaBody::Block(stmts) = body {
                    crate::compiler::collect_declared_names(stmts, &mut local_decls);
                }
                for (idx, cap_name) in captured_names.iter().enumerate() {
                    if let Some(param_slot) = self.scope().resolve(cap_name) {
                        self.emit_u16(Op::LOCAL_GET, param_slot);
                        crate::emitter::closures::emit_env_set(
                            self.chunk(),
                            env_slot,
                            idx as u16,
                            line,
                        );
                    } else if !local_decls.contains(cap_name) && parent_shared_env_slot.is_some() {
                        if let Some(parent_idx) =
                            parent_shared_env_names.iter().position(|n| n == cap_name)
                        {
                            let closure_env = self.closure_env_slot();
                            crate::emitter::closures::emit_env_get(
                                self.chunk(),
                                closure_env,
                                parent_idx as u16,
                                line,
                            );
                            crate::emitter::closures::emit_env_set(
                                self.chunk(),
                                env_slot,
                                idx as u16,
                                line,
                            );
                        }
                    }
                }
            }
        }

        let result_slot = if self.profile.function_return == ReturnStyle::ResultSlot {
            let rs = self.define_local("Result");
            self.emit(Op::NULL);
            self.emit_u16(Op::LOCAL_SET, rs);
            let saved_rs = self.current_result_slot.take();
            self.current_result_slot = Some(rs);
            Some((rs, saved_rs))
        } else {
            None
        };
        let saved_result_slot = result_slot.as_ref().map(|(_, saved_rs)| *saved_rs);

        let async_try = if is_async && self.is_js_profile() {
            let line = self.line;
            Some(common::functions::emit_async_body_start(
                &mut self.chunks[self.current],
                line,
            ))
        } else {
            None
        };
        if async_try.is_some() {
            self.active_async_try_depth += 1;
        }

        match body {
            LambdaBody::Expr(expr) => {
                self.compile_expr(expr)?;
                if self.current_chunk_is_js_async() {
                    let resolve_idx = self.import("ecma:promise", "resolve");
                    self.emit_host_call(resolve_idx, 1);
                    self.emit_return_through_finally(1)?;
                } else {
                    self.emit(Op::RETURN);
                }
            }
            LambdaBody::Block(stmts) => {
                for s in stmts {
                    self.compile_stmt(s)?;
                }
            }
        }

        if async_try.is_some() {
            self.active_async_try_depth = self.active_async_try_depth.saturating_sub(1);
        }

        if let Some(catch_jump) = async_try {
            let line = self.line;
            let chunk = &mut self.chunks[self.current];
            common::functions::emit_async_body_fallthrough(chunk, catch_jump, line);
            let resolve_idx = self.import("ecma:promise", "resolve");
            self.emit_host_call(resolve_idx, 1);
            self.emit(Op::RETURN);
            let chunk = &mut self.chunks[self.current];
            common::functions::patch_async_body_catch(chunk, catch_jump);
            let reject_idx = self.import("ecma:promise", "reject");
            self.emit_host_call(reject_idx, 1);
            self.emit(Op::RETURN);
        } else if let Some((rs, saved_rs)) = result_slot {
            self.emit_u16(Op::LOCAL_GET, rs);
            self.emit(Op::RETURN);
            self.current_result_slot = saved_rs;
        } else if matches!(body, LambdaBody::Block(_)) {
            let line = self.line;
            common::functions::emit_function_epilogue(&mut self.chunks[ci], line);
        }
        if let Some(saved_rs) = saved_result_slot {
            self.current_result_slot = saved_rs;
        }

        self.current_func_name = saved_fn;
        self.in_strict = saved_strict;

        let ns = self.scope().next_slot;
        self.chunks[ci].finalize_local_count(ns);
        let uvs = self.scopes.last().unwrap().upvalues.clone();
        // Resolve upvalue names BEFORE popping the inner scope
        let inner_scope_idx = self.scopes.len() - 1;
        let uv_names: Vec<Option<String>> = (0..uvs.len())
            .map(|i| self.captured_name_for_upvalue(inner_scope_idx, i as u8))
            .collect();
        self.scopes.pop();
        self.current = saved;
        self.active_async_try_depth = saved_async_try_depth;
        self.current_closure_captured_locals = saved_closure_captured;
        self.closure_env_names = saved_env_names;
        self.capture_locals = saved_capture_locals;
        self.shared_env_slot = saved_shared_env_slot;
        self.shared_env_names = saved_shared_env_names;
        let parent_locals = self.scope().locals.clone();
        let line = self.line;
        if uvs.is_empty() {
            common::functions::emit_ref_func(&mut self.chunks[self.current], ci, 0, line);
        } else if let Some(shared_slot) = parent_shared_env_slot {
            // Parent has a shared env — pass it directly as the upvalue.
            // The inner function's closure_env_names was pre-seeded from
            // parent_shared_env_names, so indices match.
            common::functions::emit_ref_func(&mut self.chunks[self.current], ci, 1, line);
            self.chunks[self.current].emit(1, line); // is_local = true
            self.chunks[self.current].emit(shared_slot as u8, line);
        } else {
            // No shared env — build a per-closure env (original path).
            let mut env_slots: Vec<u16> = Vec::new();
            let mut env_names: Vec<String> = Vec::new();
            for (i, uv) in uvs.iter().enumerate() {
                if let Some(name) = uv_names[i].clone() {
                    let slot = if uv.is_local {
                        let by_value = parent_locals
                            .iter()
                            .find(|l| l.slot == uv.index as u16)
                            .map(|l| {
                                self.capture_by_value_vars
                                    .iter()
                                    .any(|n| *n == self.canon(&l.name))
                            })
                            .unwrap_or(false);
                        if by_value {
                            let orig_slot = uv.index as u16;
                            self.emit_u16(Op::LOCAL_GET, orig_slot);
                            let snap = self.define_local(&format!("__snap_{}_{}", name, ci));
                            self.emit_u16(Op::LOCAL_SET, snap);
                            snap
                        } else {
                            uv.index as u16
                        }
                    } else {
                        let parent_env = self.closure_env_slot();
                        let parent_idx = self.closure_env_index(&name);
                        let tmp = self.define_local(&format!("__nested_cap_{}", name));
                        crate::emitter::closures::emit_env_get(
                            self.chunk(),
                            parent_env,
                            parent_idx,
                            line,
                        );
                        self.emit_u16(Op::LOCAL_SET, tmp);
                        tmp
                    };
                    env_names.push(name);
                    env_slots.push(slot);
                }
            }
            crate::emitter::closures::emit_env_new(self.chunk(), &env_slots, line);
            let env_slot = self.define_local(&format!("__closure_env_{}", ci));
            self.emit_u16(Op::LOCAL_SET, env_slot);
            common::functions::emit_ref_func(&mut self.chunks[self.current], ci, 1, line);
            self.chunks[self.current].emit(1, line); // is_local = true
            self.chunks[self.current].emit(env_slot as u8, line);
        }
        if self.is_js_profile() {
            let length = params
                .iter()
                .take_while(|p| p.default.is_none() && !p.is_rest)
                .count();

            inst!(self, core_wasm::dup);
            self.emit_const(Value::F64(length as f64));
            let length_key = self.str_const("length");
            self.emit_u16(Op::STRUCT_SET, length_key);
            self.emit(Op::DROP);

            inst!(self, core_wasm::dup);
            {
                // Recover the source kind when the walker lowered a generator
                // into a plain wrapper holding `__gen_fn` (obj-literal
                // `*m(){}` methods and generator expressions).
                let (eff_async, eff_generator) = match body {
                    LambdaBody::Block(stmts) => Self::wrapped_generator_kind(stmts)
                        .unwrap_or((is_async, is_generator)),
                    _ => (is_async, is_generator),
                };
                let line = self.line;
                crate::emitter::prototypes::emit_stamp_function_kind_proto(
                    self.chunk(),
                    eff_async,
                    eff_generator,
                    line,
                );
            }

            // §7.2.4 IsConstructor: arrows (and the other Lambda-lowered
            // forms — shorthand methods, generator expressions) have no
            // [[Construct]]. `new` on them must TypeError; the host
            // construct path checks this marker.
            inst!(self, core_wasm::dup);
            self.emit_const(Value::Bool(true));
            let non_ctor_key = self.str_const("__vybe_non_ctor");
            self.emit_u16(Op::STRUCT_SET, non_ctor_key);
            self.emit(Op::DROP);

            // §10.2.9/§10.2.10: name/length are non-enumerable.
            inst!(self, core_wasm::dup);
            {
                let line = self.line;
                crate::emitter::prototypes::emit_stamp_fn_metadata_nonenum(self.chunk(), line);
            }

            // §10.2.11: arrows carry a marker — the host uses it for
            // lexical-this (call/apply ignore thisArg) and toString's
            // `=>` form. Threaded from the ExprKind::Lambda compile arm;
            // object-literal shorthand methods pass false.
            if is_arrow {
                inst!(self, core_wasm::dup);
                self.emit_const(Value::Bool(true));
                let arrow_key = self.str_const("__fn_arrow");
                self.emit_u16(Op::STRUCT_SET, arrow_key);
                self.emit(Op::DROP);
            }
        }
        if has_rest {
            self.emit_stamp_rest_metadata_on_stack(params.len().saturating_sub(1));
        }
        Ok(())
    }

    /// ES2024 `Object.groupBy(arr, fn)` — inline loop emitter.
    ///
    /// Stack on entry: [arr, fn]. Result: new object whose keys are the
    /// string results of `fn(item)` and whose values are arrays of matching
    /// items. Uses only already-registered host fns (ecma:object, ecma:array);
    /// no new imports needed.
    pub(super) fn emit_object_group_by(&mut self, line: u32) -> Result<(), String> {
        let fn_slot = self.define_local("__groupby_fn");
        self.emit_u16(Op::LOCAL_SET, fn_slot);
        let arr_slot = self.define_local("__groupby_arr");
        self.emit_u16(Op::LOCAL_SET, arr_slot);

        let new_idx = self.import("ecma:object", "new");
        self.emit_host_call(new_idx, 0);
        let result_slot = self.define_local("__groupby_result");
        self.emit_u16(Op::LOCAL_SET, result_slot);

        self.emit_u16(Op::LOCAL_GET, arr_slot);
        common::collections::emit_len(&mut self.chunks, self.current, line);
        let len_slot = self.define_local("__groupby_len");
        self.emit_u16(Op::LOCAL_SET, len_slot);

        self.emit_const(Value::F64(0.0));
        let i_slot = self.define_local("__groupby_i");
        self.emit_u16(Op::LOCAL_SET, i_slot);

        let loop_state = common::loops::emit_loop_start(&mut self.chunks, self.current, line);

        self.emit_u16(Op::LOCAL_GET, i_slot);
        self.emit_u16(Op::LOCAL_GET, len_slot);
        {
            let line = self.line;
            crate::emitter::ops::emit_dyn_lt(self.chunk(), line);
        };
        common::loops::emit_loop_cond(&mut self.chunks, self.current, line);

        self.emit_u16(Op::LOCAL_GET, arr_slot);
        self.emit_u16(Op::LOCAL_GET, i_slot);
        common::collections::emit_get(&mut self.chunks, self.current, line);
        let item_slot = self.define_local("__groupby_item");
        self.emit_u16(Op::LOCAL_SET, item_slot);

        // key = fn(item)
        self.emit_u16(Op::LOCAL_GET, fn_slot);
        self.emit_u16(Op::LOCAL_GET, item_slot);
        self.emit_u8(Op::CALL_REF, 1);
        let key_slot = self.define_local("__groupby_key");
        self.emit_u16(Op::LOCAL_SET, key_slot);

        // if result[key] === undefined: result[key] = []
        self.emit_u16(Op::LOCAL_GET, result_slot);
        self.emit_u16(Op::LOCAL_GET, key_slot);
        common::collections::emit_get(&mut self.chunks, self.current, line);
        fn_call!(self, "wasm:js-undefined", "test", 1);
        let line = self.line;
        self.chunk().emit_if(line);
        self.emit_u16(Op::LOCAL_GET, result_slot);
        self.emit_u16(Op::LOCAL_GET, key_slot);
        common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
        common::collections::emit_set(&mut self.chunks, self.current, line);
        self.emit(Op::DROP);
        self.chunk().emit_end(line);

        // result[key].push(item)
        self.emit_u16(Op::LOCAL_GET, result_slot);
        self.emit_u16(Op::LOCAL_GET, key_slot);
        common::collections::emit_get(&mut self.chunks, self.current, line);
        self.emit_u16(Op::LOCAL_GET, item_slot);
        common::collections::emit_push(&mut self.chunks, self.current, line);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, i_slot);
        self.emit_const(Value::F64(1.0));
        {
            let line = self.line;
            crate::emitter::ops::emit_dyn_add(self.chunk(), line);
        };
        self.emit_u16(Op::LOCAL_SET, i_slot);

        common::loops::emit_loop_end(&mut self.chunks, self.current, loop_state, line);
        self.emit_u16(Op::LOCAL_GET, result_slot);
        Ok(())
    }

    fn try_compile_dotnet_component_call(
        &mut self,
        parts: &[String],
        args: &[&Expression],
    ) -> Result<bool, String> {
        let lower_parts: Vec<String> = parts.iter().map(|s| s.to_lowercase()).collect();
        let refs: Vec<&str> = lower_parts.iter().map(|s| s.as_str()).collect();
        let imports = crate::platforms::dotnet::emitter::core::imports::default_interface_imports();
        let ctx = common::dotnet::ResolutionContext {
            is_local: &|_: &str| false,
            is_class_field: &|_: &str| false,
            is_user_type: &|_: &str| false,
            imports: &imports,
        };
        let resolution = common::dotnet::resolve_dotted_name(&refs, &ctx);
        match resolution {
            common::dotnet::DottedResolution::CommonCall { emit } => {
                for a in args {
                    self.compile_expr(a)?;
                }
                let line = self.line;
                self.emit_common(&emit, args.len() as u8, line);
                Ok(true)
            }
            common::dotnet::DottedResolution::HostCall { module, func } => {
                for a in args {
                    self.compile_expr(a)?;
                }
                let idx = self.import(&module, &func);
                self.emit_host_call(idx, args.len() as u8);
                Ok(true)
            }
            _ => Ok(false),
        }
    }
}
