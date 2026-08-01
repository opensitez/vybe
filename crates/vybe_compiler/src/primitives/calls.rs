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

pub(super) fn terminal_type_name(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => Some(name.rsplit('.').next().unwrap_or(name).to_string()),
        ExprKind::Member { field, .. } => Some(field.clone()),
        _ => None,
    }
}

pub(super) fn strip_generic_suffix(name: &str) -> &str {
    common::generics::generic_base_name(name)
}

pub(super) fn extract_generic_type_name(name: &str) -> Option<String> {
    common::generics::first_generic_argument_leaf_name(name)
}

fn array_element_type_hint(type_hint: &str) -> Option<String> {
    let trimmed = type_hint.trim();
    if let Some(element) = trimmed.strip_suffix("()") {
        return Some(element.trim().to_string());
    }
    if trimmed.ends_with(']') {
        let bracket = trimmed.rfind('[')?;
        return Some(trimmed[..bracket].trim().to_string());
    }
    None
}

/// Borrow a configured namespace path as the `&[&str]` the tree walk takes.
fn scope_segments(scope: &[String]) -> Vec<&str> {
    scope.iter().map(String::as_str).collect()
}

fn is_dotnet_linq_method_name(name: &str) -> bool {
    matches!(
        name.to_ascii_lowercase().as_str(),
        "aggregate"
            | "all"
            | "any"
            | "append"
            | "asenumerable"
            | "average"
            | "cast"
            | "chunk"
            | "concat"
            | "contains"
            | "count"
            | "defaultifempty"
            | "distinct"
            | "distinctby"
            | "elementat"
            | "elementatordefault"
            | "except"
            | "exceptby"
            | "first"
            | "firstordefault"
            | "groupby"
            | "intersect"
            | "intersectby"
            | "last"
            | "lastordefault"
            | "longcount"
            | "max"
            | "maxby"
            | "min"
            | "minby"
            | "oftype"
            | "orderby"
            | "orderbydescending"
            | "prepend"
            | "reverse"
            | "select"
            | "selectmany"
            | "sequenceequal"
            | "single"
            | "singleordefault"
            | "skip"
            | "skiplast"
            | "skipwhile"
            | "sum"
            | "take"
            | "takelast"
            | "takewhile"
            | "thenby"
            | "thenbydescending"
            | "toarray"
            | "todictionary"
            | "tolist"
            | "tolookup"
            | "union"
            | "unionby"
            | "where"
            | "zip"
    )
}

fn dotnet_factory_return_type(scope: &[String], callee: &Expression) -> Option<String> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let class_name = terminal_type_name(object)?;
    vybe_runtime::namespaces::lookup_type_member_return(scope, &class_name, field)
}

fn dotnet_static_member_return_type(scope: &[String], expr: &Expression) -> Option<String> {
    let ExprKind::Member { object, field, .. } = &expr.kind else {
        return None;
    };
    let class_name = terminal_type_name(object)?;
    vybe_runtime::namespaces::lookup_type_member_return(scope, &class_name, field)
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

pub(super) fn resolve_receiver_type_hint(compiler: &Compiler, recv: &Expression) -> Option<String> {
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
            if let Some(type_name) =
                dotnet_static_member_return_type(&compiler.profile.namespaces.type_scopes, recv)
            {
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
        ExprKind::Index { object, .. } if compiler.profile.namespaces.use_dotnet => {
            resolve_receiver_type_hint(compiler, object)
                .as_deref()
                .and_then(array_element_type_hint)
                .map(|name| compiler.resolve_source_type_alias(&name))
        }
        ExprKind::Call { callee, args, .. } => {
            if compiler.profile.parens_for_index && args.len() == 1 {
                if let Some(type_hint) = resolve_receiver_type_hint(compiler, callee) {
                    if let Some(element_type) = array_element_type_hint(&type_hint) {
                        return Some(compiler.resolve_source_type_alias(&element_type));
                    }
                }
            }
            let arg_exprs: Vec<&Expression> = args.iter().map(|arg| &arg.value).collect();
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                if let Some(return_type) = compiler
                    .resolve_instance_method_overload(object, field, &arg_exprs, false)
                    .and_then(|overload| overload.return_type.clone())
                {
                    return Some(return_type);
                }

                if !compiler.profile.namespaces.type_scopes.is_empty() {
                    if let Some(receiver_type) = resolve_receiver_type_hint(compiler, object) {
                        if compiler
                            .resolve_pending_class_name_for_type_hint(&receiver_type)
                            .is_none()
                        {
                            let class_name = Compiler::normalize_type_hint(&receiver_type);
                            if let Some(return_type) =
                                vybe_runtime::namespaces::lookup_type_member_return(
                                    &compiler.profile.namespaces.type_scopes,
                                    &class_name,
                                    field,
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
                .or_else(|| {
                    dotnet_factory_return_type(&compiler.profile.namespaces.type_scopes, callee)
                })
                .or_else(|| match &callee.kind {
                    ExprKind::Ident(name) => {
                        let resolved = compiler.resolve_source_type_alias(name);
                        vybe_runtime::namespaces::lookup_type_ctor_target(
                            &compiler.profile.namespaces.type_scopes,
                            &resolved,
                        )
                        .map(|_| resolved)
                    }
                    ExprKind::Member { field, .. } => {
                        let resolved = compiler.resolve_source_type_alias(field);
                        vybe_runtime::namespaces::lookup_type_ctor_target(
                            &compiler.profile.namespaces.type_scopes,
                            &resolved,
                        )
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
        // An array literal receiver (`new[]{1,2,3}.Where(...)`, `{1,2}.Sum()`)
        // is an `IEnumerable<T>` in .NET, so LINQ resolves against the shared
        // surface. Gated on `use_dotnet` so non-.NET languages (Ruby `.select`,
        // JS array HOFs) keep their own array-method semantics.
        ExprKind::Array(_) if compiler.profile.namespaces.use_dotnet => {
            Some("IEnumerable".to_string())
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

        let ctor_global = crate::primitives::classes::ctor_global_for(&class_name, 0);
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
        self.normalized_overload_type_matches(&normalized_param, &normalized_arg)
    }

    fn overload_type_exact_matches(&self, param_type: &str, arg_type: &str) -> bool {
        let normalized_param =
            Self::normalize_type_hint(strip_generic_suffix(param_type).trim_end_matches('?'));
        let normalized_arg =
            Self::normalize_type_hint(strip_generic_suffix(arg_type).trim_end_matches('?'));
        normalized_param == normalized_arg
            || (Self::is_string_type_hint(&normalized_param)
                && Self::is_string_type_hint(&normalized_arg))
            || (matches!(normalized_param.as_str(), "bool" | "boolean")
                && matches!(normalized_arg.as_str(), "bool" | "boolean"))
    }

    fn normalized_overload_type_matches(
        &self,
        normalized_param: &str,
        normalized_arg: &str,
    ) -> bool {
        normalized_param == normalized_arg
            || (Self::is_string_type_hint(&normalized_param)
                && Self::is_string_type_hint(&normalized_arg))
            || (matches!(normalized_param, "bool" | "boolean")
                && matches!(normalized_arg, "bool" | "boolean"))
            || (is_numeric_overload_type(normalized_param)
                && is_numeric_overload_type(normalized_arg))
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
        let resolved_receiver_type = self.resolve_source_type_alias(receiver_type);
        let receiver_canon = self
            .canon(strip_generic_suffix(&resolved_receiver_type))
            .replace('\\', ".");
        if self.pending_classes.contains_key(&receiver_canon) {
            return Some(receiver_canon);
        }
        let resolved_canon = self.canon(&resolved_receiver_type).replace('\\', ".");
        if self.pending_classes.contains_key(&resolved_canon) {
            return Some(resolved_canon);
        }
        let receiver_canon_raw = self.canon(receiver_type).replace('\\', ".");
        if self.pending_classes.contains_key(&receiver_canon_raw) {
            return Some(receiver_canon_raw);
        }

        let mut matches = self.pending_classes.keys().filter(|name| {
            let simple_name = name.rsplit('.').next().unwrap_or(name);
            name.eq_ignore_ascii_case(receiver_type)
                || name.eq_ignore_ascii_case(&resolved_receiver_type)
                || name.eq_ignore_ascii_case(&receiver_canon)
                || simple_name.eq_ignore_ascii_case(&receiver_canon)
        });
        match (matches.next(), matches.next()) {
            (Some(name), None) => Some(name.clone()),
            _ => None,
        }
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

    pub(super) fn overload_storage_name(
        &self,
        method_name: &str,
        param_types: &[String],
    ) -> String {
        if param_types.is_empty() {
            format!("{method_name}$sig0")
        } else {
            format!("{method_name}$sig{}", param_types.join("$"))
        }
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

        for exact_only in [true, false] {
            'overload_search: for overload in overloads {
                let signature = &overload.signature;
                let param_count = overload.param_types.len();
                let arity_ok = actual_arity >= signature.min_arity
                    && (signature.has_rest || actual_arity <= param_count);
                if !arity_ok {
                    continue;
                }

                for (arg_expr, param_type) in effective_args.iter().zip(overload.param_types.iter())
                {
                    if let Some(arg_type) = self.infer_expr_type_hint(arg_expr) {
                        let matches = if exact_only {
                            self.overload_type_exact_matches(param_type, &arg_type)
                        } else {
                            self.overload_type_matches(param_type, &arg_type)
                        };
                        if !matches {
                            continue 'overload_search;
                        }
                    }
                }

                return Some(overload.clone());
            }
        }

        None
    }

    /// The chunk to bind a member call to DIRECTLY, or `None` to fall through
    /// to the dynamic member lookup.
    ///
    /// The overload is selected from the receiver's declared type, which is
    /// correct — overload selection is a compile-time decision from the static
    /// argument types in every language that has overloads. Picking *which
    /// body runs* is not: for a virtual method the declared type's chunk is
    /// the wrong target whenever the runtime type overrides it
    /// (`Base b = new Derived(); b.Speak()`). Declining the direct bind there
    /// falls through to the dynamic path, which already resolves overrides
    /// correctly for untyped receivers and casts.
    ///
    /// One guard keeps the decline safe because the dynamic path resolves by
    /// member NAME alone:
    ///
    /// - Only single-overload names may decline. An overloaded name has one
    ///   runtime slot, so declining would silently pick the last-registered
    ///   signature and turn a dispatch bug into an overload bug. An overloaded
    ///   virtual keeps its declared-type bind: no better, but no worse.
    ///
    /// Hidden methods (C# `new`, VB `Shadows`) bind into class-qualified
    /// storage slots, so they do not overwrite the virtual slot and no longer
    /// need to suppress dynamic dispatch here.
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

    /// Resolve an instance method on a typed receiver through the shared
    /// namespace tree: walk user pending-classes into the first registered
    /// platform type. A user-declared member of the same name is an override
    /// and stays dynamic (`None`); otherwise the registered type owns the
    /// method target.
    pub(super) fn namespace_tree_instance_method_owner(
        &self,
        type_hint: &str,
        method_name: &str,
        arg_count: u8,
    ) -> Option<String> {
        let scope = &self.profile.namespaces.type_scopes;
        let mut current = self
            .resolve_pending_class_name_for_type_hint(type_hint)
            .unwrap_or_else(|| Self::normalize_type_hint(type_hint));
        loop {
            if vybe_runtime::namespaces::is_registered_type(scope, &current) {
                // Platform class — its registered members finish the chain.
                return vybe_runtime::namespaces::lookup_type_instance_target(
                    scope,
                    &current,
                    method_name,
                    arg_count,
                )
                .map(|_| current);
            }
            let pending = self.pending_classes.get(&current)?;
            let key = self.js_member_storage_name_for_class(&current, method_name);
            if pending
                .instance_member_names
                .iter()
                .any(|name| name == &key)
            {
                return None;
            }
            current = pending.parent.clone()?;
        }
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

    pub(crate) fn resolve_unique_static_method_chunk_for_class(
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
        self.emit_instance_method_call_from_fn_slot(fn_tmp, method_name, obj_tmp, args, arg_exprs)
    }

    fn emit_instance_method_call_from_fn_slot(
        &mut self,
        fn_tmp: u16,
        method_name: &str,
        obj_tmp: u16,
        args: &[Argument],
        arg_exprs: &[&Expression],
    ) -> Result<(), String> {
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

                if self.class_prototype_dispatch() {
                    self.emit_call_ref_with_bound_js_this_arg_slots(fn_tmp, obj_tmp, &arg_slots);
                } else {
                    self.emit_call_ref_with_arg_slots(fn_tmp, Some(obj_tmp), &arg_slots);
                }

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
        if self.class_prototype_dispatch() {
            self.emit_call_ref_with_bound_js_this_arg_slots(fn_tmp, obj_tmp, &arg_slots);
        } else {
            self.emit_call_ref_with_arg_slots(fn_tmp, Some(obj_tmp), &arg_slots);
        }
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
        if !self.profile.ambient_this_binding {
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
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
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
        self.stamp_lua_multi_row_slot(rest_slot);
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
            crate::primitives::ops::emit_dyn_gt(self.chunk(), line);
        };
        let line = self.line;
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
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
        if self.uses_proxy && receiver_slot.is_none() {
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
                crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
            };
            let line = self.line;
            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
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
        if let Some(receiver_slot) = receiver_slot {
            if self.class_prototype_dispatch() {
                let result_slot = self.define_local("__call_runtime_result");
                let has_own_marker_slot =
                    self.emit_js_has_own_receiver_marker(callee_slot, "__js_receiver_call_marker");
                let receiver_key = self.str_const("__vybe_method_receiver");
                self.emit_u16(Op::LOCAL_GET, callee_slot);
                self.emit_u16(Op::STRUCT_GET, receiver_key);
                let marker_slot = self.define_local("__js_receiver_call_marker_value");
                self.emit_u16(Op::LOCAL_SET, marker_slot);

                self.emit_u16(Op::LOCAL_GET, has_own_marker_slot);
                self.emit(Op::I32_EQZ);
                let line = self.line;
                self.chunk().emit_if(line);
                self.emit_dispatch_and_store_from_arg_slots(
                    callee_slot,
                    None,
                    Some(receiver_slot),
                    arg_slots,
                    result_slot,
                );
                self.chunk().emit_else(line);
                self.emit_u16(Op::LOCAL_GET, marker_slot);
                fn_call!(self, "wasm:js-undefined", "test", 1);
                let line = self.line;
                self.chunk().emit_if(line);
                self.emit_dispatch_and_store_from_arg_slots(
                    callee_slot,
                    None,
                    Some(receiver_slot),
                    arg_slots,
                    result_slot,
                );
                self.chunk().emit_else(line);
                self.emit_js_apply_from_arg_slots(callee_slot, receiver_slot, arg_slots);
                self.emit_u16(Op::LOCAL_SET, result_slot);
                self.chunk().emit_end(line);
                self.chunk().emit_end(line);
                self.emit_u16(Op::LOCAL_GET, result_slot);
                return;
            }
        }
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

    fn emit_js_receiver_host_or_bound_this_call(
        &mut self,
        callee_slot: u16,
        receiver_slot: u16,
        arg_slots: &[u16],
    ) {
        let has_own_marker_slot =
            self.emit_js_has_own_receiver_marker(callee_slot, "__js_receiver_host_marker");
        let receiver_key = self.str_const("__vybe_method_receiver");
        self.emit_u16(Op::LOCAL_GET, callee_slot);
        self.emit_u16(Op::STRUCT_GET, receiver_key);
        let marker_slot = self.define_local("__js_receiver_host_marker_value");
        self.emit_u16(Op::LOCAL_SET, marker_slot);

        self.emit_u16(Op::LOCAL_GET, has_own_marker_slot);
        self.emit(Op::I32_EQZ);
        let line = self.line;
        self.chunk().emit_if(line);
        self.emit_call_ref_with_bound_js_this_arg_slots(callee_slot, receiver_slot, arg_slots);
        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, marker_slot);
        fn_call!(self, "wasm:js-undefined", "test", 1);
        let line = self.line;
        self.chunk().emit_if(line);
        self.emit_call_ref_with_bound_js_this_arg_slots(callee_slot, receiver_slot, arg_slots);
        self.chunk().emit_else(line);
        self.emit_js_apply_from_arg_slots(callee_slot, receiver_slot, arg_slots);
        self.chunk().emit_end(line);
        self.chunk().emit_end(line);
    }

    fn emit_js_has_own_receiver_marker(&mut self, callee_slot: u16, local_name: &str) -> u16 {
        self.emit_u16(Op::LOCAL_GET, callee_slot);
        self.emit_const(Value::String(Arc::from("__vybe_method_receiver")));
        let has_own_idx = self.import("ecma:object", "hasOwn");
        self.emit_host_call(has_own_idx, 2);
        let line = self.line;
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        let slot = self.define_local(local_name);
        self.emit_u16(Op::LOCAL_SET, slot);
        slot
    }

    fn emit_js_apply_from_arg_slots(
        &mut self,
        callee_slot: u16,
        receiver_slot: u16,
        arg_slots: &[u16],
    ) {
        self.emit_u16(Op::LOCAL_GET, callee_slot);
        self.emit_u16(Op::LOCAL_GET, receiver_slot);
        let args_slot = self.define_local("__js_apply_args");
        common::collections::emit_array_new(&mut self.chunks, self.current, 0, self.line);
        self.emit_u16(Op::LOCAL_SET, args_slot);
        for slot in arg_slots {
            self.emit_u16(Op::LOCAL_GET, args_slot);
            self.emit_u16(Op::LOCAL_GET, *slot);
            common::collections::emit_push(&mut self.chunks, self.current, self.line);
            self.emit(Op::DROP);
        }
        self.emit_u16(Op::LOCAL_GET, args_slot);
        fn_call!(self, "ecma:function", "apply", 3);
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
    ) -> Result<(), String> {
        // A method call on a null/undefined receiver is a *catchable* error,
        // not a VM trap. Throw through the common `errors.rs` machinery; the
        // exception type is language-defined (JS `TypeError`, PHP `Error`, …)
        // and read from the profile so this stays language-agnostic and
        // cross-language catchable.
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit(Op::REF_IS_NULL);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        fn_call!(self, "wasm:js-undefined", "test", 1);
        self.emit(Op::I32_OR);
        let nline = self.line;
        self.chunk().emit_if(nline);
        let msg = format!("Call to a member function {}() on null", method_name);
        self.emit_const(Value::String(Arc::from(msg.as_str())));
        let err_type = self.profile.member_call_on_null_error.clone();
        self.emit_js_exception_ctor_from_message_value(&err_type)?;
        common::errors::emit_throw(self.chunk(), nline);
        self.chunk().emit_end(nline);

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
        Ok(())
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
                    crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                };
                let line = self.line;
                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
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
        if !self.profile.supports_spread_arguments || !args.iter().any(|a| a.spread) {
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
        let line = self.line;
        common::reflection::emit_reflect_op(
            &mut self.chunks,
            self.current,
            common::reflection::ReflectOp::Apply,
            3,
            line,
        );
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
        self.chunks[func_idx].local_names = self.scope().defined_names.clone();
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

    pub(super) fn error_instanceof_chain(type_name: &str) -> &'static [&'static str] {
        match type_name.trim() {
            "Error" => &["Error"],
            "EvalError" => &["EvalError", "Error"],
            "RangeError" => &["RangeError", "Error"],
            "ReferenceError" => &["ReferenceError", "Error"],
            "SyntaxError" => &["SyntaxError", "Error"],
            "TypeError" => &["TypeError", "Error"],
            "URIError" => &["URIError", "Error"],
            "AggregateError" => &["AggregateError", "Error"],
            "SuppressedError" => &["SuppressedError", "Error"],
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

        // Canonical ancestor chain (`except LookupError:` catching KeyError).
        // NOT for `throwable_is_root` profiles (PHP/Java): their
        // Error/Exception branches are siblings and their constructors stamp
        // their own `__types`.
        if !self.profile.throwable_is_root {
            common::errors::emit_stamp_exception_ancestors(self.chunk(), type_name, line);
        }

        let exc_tmp = self.define_local("__exc_tmp");
        self.emit_u16(Op::LOCAL_SET, exc_tmp);

        if self.profile.ecma_error_object_shape {
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
                // §20.5 instance property descriptors are
                // { writable: true, enumerable: false, configurable: true } —
                // omitting these lets defineProperty default them to false,
                // which (a) violates the spec shape and (b) makes the own
                // `name` stamp undeletable for the prototype-chain finisher.
                for flag in ["writable", "configurable"] {
                    inst!(self, core_wasm::dup);
                    self.emit_const(Value::Bool(true));
                    let flag_key = self.str_const(flag);
                    self.emit_u16(Op::STRUCT_SET, flag_key);
                    self.emit(Op::DROP);
                }
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

        if self.profile.ecma_error_object_shape {
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

        if self.profile.ecma_error_object_shape {
            for name in Self::error_instanceof_chain(type_name) {
                crate::primitives::reflection::emit_instanceof_chain(
                    &mut self.chunks,
                    self.current,
                    exc_tmp,
                    name,
                    line,
                );
            }
        }

        self.emit_u16(Op::LOCAL_GET, exc_tmp);
        if self.profile.ecma_error_object_shape {
            // §20.5: link [[Prototype]] to the prelude-wired
            // `__ctor_<Kind>.prototype` and drop the own `name` stamp —
            // instances resolve `name`/`toString` through the chain.
            crate::primitives::errors::emit_finish_js_error_instance(self.chunk(), type_name, line);
        }
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

        if type_name.trim() == "SuppressedError" {
            let idx = self.import("ecma:error", "SuppressedError");
            self.emit_u16(Op::STRUCT_NEW, 0);
            for arg in args {
                self.compile_expr(arg)?;
            }
            self.emit_host_call(idx, (args.len() + 1) as u8);
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
            // `**kwargs` collector, possibly alongside a non-last `*args`. This
            // is the unified named+variadic+kwargs binding: positionals fill the
            // fixed params then the variadic collector; named args matching no
            // fixed param go into the kwargs dict. Emitted as
            // `[fixed…, rest_array?, kwargs_dict]`, which the callee binds
            // positionally — a `*a, **k` shape is never runtime-rest-packed (its
            // last param is `**k`, not the rest), so there's no double packing.
            // Data-driven via `has_kwargs`; a no-op for languages without it.
            if signature.has_kwargs && !signature.has_rest {
                let kw_index = signature.param_names.len().saturating_sub(1);
                let rest_index = signature.rest_index.filter(|&i| i < kw_index);

                let mut fixed_slots: Vec<Option<Argument>> = vec![None; kw_index];
                let mut rest_items: Vec<Expression> = Vec::new();
                let mut kwargs_entries: Vec<ObjectProperty> = Vec::new();
                let mut next_positional = 0usize;
                let mut valid = true;

                for arg in args {
                    if arg.spread {
                        valid = false;
                        break;
                    }
                    if let Some(name) = arg.name.as_deref() {
                        match signature.param_names[..kw_index]
                            .iter()
                            .position(|param_name| param_name.eq_ignore_ascii_case(name))
                        {
                            // A named arg for a real fixed param (not the variadic).
                            Some(index) if Some(index) != rest_index => {
                                if fixed_slots[index].is_some() {
                                    valid = false;
                                    break;
                                }
                                let mut ordered = arg.clone();
                                ordered.name = None;
                                fixed_slots[index] = Some(ordered);
                            }
                            // Matches the variadic's name or nothing → kwargs.
                            _ => kwargs_entries.push(ObjectProperty::KeyValue {
                                key: Expression::new(ExprKind::Lit(Literal::Str(name.to_string()))),
                                value: arg.value.clone(),
                            }),
                        }
                    } else if rest_index.is_some_and(|r| next_positional >= r) {
                        // Positional past the variadic's slot → into the rest array.
                        rest_items.push(arg.value.clone());
                    } else {
                        while next_positional < fixed_slots.len()
                            && fixed_slots[next_positional].is_some()
                        {
                            next_positional += 1;
                        }
                        if next_positional >= fixed_slots.len() {
                            valid = false;
                            break;
                        }
                        fixed_slots[next_positional] = Some(arg.clone());
                        next_positional += 1;
                    }
                }

                if !valid {
                    continue;
                }
                if fixed_slots
                    .iter()
                    .take(signature.min_arity)
                    .any(Option::is_none)
                {
                    continue;
                }

                let mut ordered_args = Vec::with_capacity(kw_index + 1);
                for (i, slot) in fixed_slots.into_iter().enumerate() {
                    if Some(i) == rest_index {
                        ordered_args.push(Argument::positional(Expression::new(ExprKind::Array(
                            rest_items
                                .iter()
                                .cloned()
                                .map(|value| ArrayElement {
                                    key: None,
                                    value,
                                    spread: false,
                                    by_ref: false,
                                })
                                .collect(),
                        ))));
                    } else {
                        match slot {
                            Some(arg) => ordered_args.push(arg),
                            None => {
                                let default =
                                    signature.param_defaults.get(i).and_then(|d| d.clone());
                                ordered_args.push(Argument::positional(
                                    default.unwrap_or_else(Expression::null),
                                ));
                            }
                        }
                    }
                }
                ordered_args.push(Argument::positional(Expression::new(ExprKind::Object(
                    kwargs_entries,
                ))));
                return ordered_args;
            }

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

            // Reassemble in positional order. A gap BEFORE the last supplied
            // arg is an omitted optional param — fill it with its declared
            // default (not `null`, which the callee would treat as an explicit
            // value and skip its default). Gaps AFTER the last supplied arg are
            // truncated, so the callee applies its own defaults exactly as it
            // would for a positional call with fewer arguments.
            let last_supplied = slots.iter().rposition(Option::is_some);
            let mut ordered_args = Vec::with_capacity(slots.len());
            for (i, slot) in slots.into_iter().enumerate() {
                match slot {
                    Some(arg) => ordered_args.push(arg),
                    None => {
                        if last_supplied.is_some_and(|last| i < last) {
                            let default = signature.param_defaults.get(i).and_then(|d| d.clone());
                            ordered_args.push(Argument::positional(
                                default.unwrap_or_else(Expression::null),
                            ));
                        }
                        // trailing gap → drop; callee fills the default
                    }
                }
            }
            return ordered_args;
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

        // First-class funcref value call (WASM `call_ref` / `call_indirect`):
        // when a local variable holds a funcref, calling it is a plain
        // `CALL_REF` — push the funcref, then the args, no receiver and no
        // method dispatch. Gated on `function_references` so no other
        // language's bare-identifier call semantics change.
        if self.profile.function_references {
            if let ExprKind::Ident(name) = &callee.kind {
                // WASM tail call: `__wasm_return_call(funcref, args…)` lowers to
                // the frame-reusing `RETURN_CALL` (spec tail-call proposal) so
                // unbounded tail recursion runs in O(1) stack. Layout matches
                // `call_value`: push the funcref, then the args.
                if name == "__wasm_return_call" && !args.is_empty() {
                    self.compile_expr(&args[0].value)?; // funcref callee
                    for a in &args[1..] {
                        self.compile_expr(&a.value)?;
                    }
                    self.emit_u8(Op::RETURN_CALL, (args.len() - 1) as u8);
                    return Ok(());
                }
                if self.scope().resolve(name).is_some() {
                    self.emit_var_get(name);
                    for a in args {
                        self.compile_expr(&a.value)?;
                    }
                    self.emit_u8(Op::CALL_REF, args.len() as u8);
                    return Ok(());
                }
            }
        }

        // Common-resolver construction (namespaceplan.md): a bare
        // `TypeName(args)` call whose name resolves — via a mounted ambient
        // root (`flutter.*`, …) — to a tree `Type` carrying a `CtorSpec` is a
        // constructor call, not a function call. Dart's `Scaffold(...)` (no
        // `new`) reaches here as a `Call`; construct it generically through the
        // ONE resolver. A local/function/user-class of the same name shadows.
        if let ExprKind::Ident(name) = &callee.kind {
            let canon = self.canon(name);
            if self.scope().resolve(name).is_none()
                && !self.defined_functions.contains(name.as_str())
                && !self.defined_functions.contains(&canon)
                && !self.defined_classes.contains(name.as_str())
                && !self.defined_classes.contains(&canon)
            {
                if let Some(super::resolver::Resolution::Tree(
                    crate::primitives::namespaces::ResolutionTarget::Ctor {
                        spec: Some(spec), ..
                    },
                )) = self.resolve_profile_namespace_chain(&[name.to_string()])
                {
                    return self.emit_tree_ctor_construction(&spec, args);
                }
            }
        }

        // A call whose callee resolves to a framework GUI control class —
        // `Form("Calculator")`, `Window.Forms.Form(...)`, `Window.Forms.TextBox()`
        // — constructs the control through the `vybe:gui` host factory, the
        // same GUI-direct path as `New Form()`. Control classes no longer emit
        // a per-class constructor global; construction resolves through the
        // component descriptor. The callee is resolved by its LAST segment
        // (namespace qualifiers are ignored, matching how the class name used
        // to resolve as a global). Guarded so a user function/class/local of
        // the same name, or a method call on an instance, is left alone.
        if self.profile.namespaces.use_dotnet {
            let parts = self.flatten_member_chain(callee);
            if let Some(last) = parts.last() {
                let canonical = common::gui::canonical_control_name(last);
                let canon_last = self.canon(last);
                let first_is_local = parts
                    .first()
                    .map_or(false, |f| self.scope().resolve(f).is_some());
                if !canonical.is_empty()
                    && !first_is_local
                    && !self.defined_functions.contains(&canon_last)
                    && !self.defined_classes.contains(&canon_last)
                {
                    let host_name = common::gui::host_fn_new_control(&canonical);
                    let new_idx = self.import("vybe:gui", &host_name);
                    for a in args {
                        self.compile_expr(&a.value)?;
                    }
                    let line = self.line;
                    common::gui::emit_new_control(self.chunk(), new_idx, args.len() as u8, line);
                    return Ok(());
                }
            }
        }

        if self.try_compile_go_map_has_call(callee, args)? {
            return Ok(());
        }

        if let ExprKind::Ident(name) = &callee.kind {
            if let Some(resolved) = self.resolve_namespaced_function_identity(name) {
                if resolved != self.canon(name) {
                    let new_callee =
                        Expression::with_span(ExprKind::Ident(resolved), callee.span.clone());
                    return self.compile_call(&new_callee, args);
                }
            }
        }

        if self.profile.supports_dynamic_import {
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
                            crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                        };
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
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
            if self.profile.namespaces.use_dotnet
                && !*null_safe
                && field.eq_ignore_ascii_case("ToString")
                && args.is_empty()
            {
                if let Some(type_hint) = self.infer_expr_type_hint(object) {
                    let resolved = self.resolve_source_type_alias(&type_hint);
                    if self
                        .namespace_tree_instance_method_owner(&resolved, field, 0)
                        .is_some()
                    {
                        // Let descriptor-owned .NET ToString implementations
                        // (StringWriter, StringBuilder, DateTime, ...) route
                        // through platforms/dotnet instead of the generic
                        // object-string fallback.
                    } else {
                        self.compile_expr(object)?;
                        let line = self.line;
                        self.emit_common("dotnet.tostring_runtime", 1, line);
                        return Ok(());
                    }
                } else {
                    self.compile_expr(object)?;
                    let line = self.line;
                    self.emit_common("dotnet.tostring_runtime", 1, line);
                    return Ok(());
                }
            }
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
        if self.try_compile_dotnet_attribute_reflection_call(callee, args)? {
            return Ok(());
        }

        if self.is_python_profile() {
            if let ExprKind::Ident(name) = &callee.kind {
                // `OrderedDict` IS a dict — ecma objects are insertion-ordered —
                // so it shares the pairs/kwargs/empty construction path (the
                // `collections` namespace tree resolves the NAME; the pairs/kwargs
                // handling needs the AST, which only this compile-time path has).
                if name == "dict" || name == "OrderedDict" {
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
                // `Counter(a=3, b=1)` — keyword form sets counts directly (needs
                // the AST, like `dict` kwargs). The positional/iterable form
                // `Counter([...])` falls through to `python.counter_new` (counting
                // loop), and empty `Counter()` too.
                if name == "Counter"
                    && !args.is_empty()
                    && args.iter().all(|arg| arg.name.is_some())
                {
                    let line = self.line;
                    common::dict::emit_new(&mut self.chunks, self.current, line);
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
            }
        }

        // `compact`/`extract` convert between a map and the set of bindings in
        // the VARIABLE namespace, so they only mean anything in a language that
        // has one. They read and write local slots BY NAME, which is why they
        // stay here rather than moving to a profile builtin: `emit_common`
        // receives a chunk and an argc, and cannot reach the scope table.
        if let Some(ns) = self.variable_namespace {
            if let ExprKind::Ident(name) = &callee.kind {
                if name == "compact" {
                    let line = self.line;
                    common::collections::emit_map_new(&mut self.chunks, self.current, line);
                    for arg in args {
                        let ExprKind::Lit(Literal::Str(var_name)) = &arg.value.kind else {
                            self.emit(Op::NULL);
                            return Ok(());
                        };
                        let var_binding = (ns.spell)(var_name);
                        inst!(self, core_wasm::dup);
                        self.emit_const(Value::String(Arc::from(var_name.as_str())));
                        self.emit_var_get(&var_binding);
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
                            let bind_body = match &key_expr.kind {
                                ExprKind::Lit(Literal::Str(s)) => s.to_string(),
                                ExprKind::Lit(Literal::Int(n)) => n.to_string(),
                                _ => continue,
                            };
                            let bind_name = (ns.spell)(&bind_body);
                            self.compile_expr(&elem.value)?;
                            self.emit_var_set(&bind_name);
                            count += 1;
                        }
                        self.emit_const(Value::I64(count));
                        return Ok(());
                    }

                    // Every USER binding in scope: a name in the variable
                    // namespace whose body is not a compiler temporary. The
                    // `__` convention for temporaries is shared (see
                    // `track_lexical_name`); the marker that makes it a
                    // variable is the language's.
                    let is_user_variable = |name: &str| {
                        self.is_variable_name(name)
                            && !self.variable_name_body(name).starts_with("__")
                    };
                    let mut binding_names = std::collections::BTreeSet::new();
                    for local in &self.scope().locals {
                        if is_user_variable(&local.name) {
                            binding_names.insert(local.name.clone());
                        }
                    }
                    for global in &self.defined_globals {
                        if is_user_variable(global) {
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
                            let key_name = self.variable_name_body(&bind_name);
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
                                crate::primitives::ops::emit_dyn_add(self.chunk(), line);
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
                            crate::primitives::classes::emit_super_once_guard(
                                self.chunk(),
                                ctx_slot,
                                l,
                            );
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
                if class_name.is_none() {
                    self.emit_const(Value::String(Arc::from("'super' keyword unexpected here")));
                    let line = self.line;
                    self.emit_js_exception_ctor_from_message_value("ReferenceError")?;
                    common::errors::emit_throw(self.chunk(), line);
                    return Ok(());
                }
                let parent_name = class_name
                    .as_ref()
                    .and_then(|cn| self.pending_classes.get(cn.as_str()))
                    .and_then(|pc| pc.parent.clone());
                let self_kw = self.profile.self_keyword.clone();
                let self_slot = self
                    .scope()
                    .resolve(&self_kw)
                    .or_else(|| self.scope().resolve_ci(&self_kw));

                if let Some(_parent) = parent_name {
                    if self.profile.class_multiple_inheritance {
                        // Cooperative super (multiple inheritance): resolve the NEXT
                        // method by walking the instance's runtime C3 MRO from the
                        // class this `super()` textually belongs to — so B.f's
                        // `super().f()` reaches C (not B's static parent A) when self
                        // is a D. `__mro__` carries the full C3 from Tier 1.
                        let cur_canon = self.canon(class_name.as_deref().unwrap_or(""));
                        let line = self.line;
                        let helper = crate::primitives::classes::ensure_super_lookup_chunk(
                            &mut self.chunks,
                            line,
                        );
                        self.emit_u16(Op::REF_FUNC, helper as u16);
                        self.chunk().emit(0, line); // 0 upvalues
                        if let Some(s) = self_slot {
                            self.emit_u16(Op::LOCAL_GET, s);
                        } else {
                            self.emit(Op::NULL);
                        }
                        self.emit_const(Value::String(Arc::from(cur_canon.as_str())));
                        self.emit_const(Value::String(Arc::from(canon_field.as_str())));
                        self.emit_u8(Op::CALL_REF, 3);
                    } else {
                        // ECMA `super` resolves from the method's
                        // [[HomeObject]].[[Prototype]] at call time. For
                        // instance methods that is `C.prototype.__proto__`,
                        // so prototype rebinding remains observable.
                        self.emit_js_super_home_base();
                        let method_idx = self.str_const(&canon_field);
                        self.emit_u16(Op::STRUCT_GET, method_idx);
                    }

                    if self.profile.ambient_this_binding {
                        let method_slot = self.define_local("__js_super_method_fn");
                        self.emit_u16(Op::LOCAL_SET, method_slot);
                        self.emit_js_current_this_value();
                        let receiver_slot = self.define_local("__js_super_method_receiver");
                        self.emit_u16(Op::LOCAL_SET, receiver_slot);
                        let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                        for (index, a) in arg_exprs.iter().enumerate() {
                            self.compile_expr(a)?;
                            let arg_slot =
                                self.define_local(&format!("__js_super_method_arg_{}", index));
                            self.emit_u16(Op::LOCAL_SET, arg_slot);
                            arg_slots.push(arg_slot);
                        }
                        self.emit_js_receiver_host_or_bound_this_call(
                            method_slot,
                            receiver_slot,
                            &arg_slots,
                        );
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
        // locals with type annotations. .NET-shaped profiles ONLY —
        // ungated, this hijacked typed receivers in other languages
        // (the "dotnet adapter leaked into compiler core" disease).
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            if self.profile.namespaces.use_dotnet
                && field.eq_ignore_ascii_case("Add")
                && arg_exprs.len() == 1
                && matches!(&object.kind, ExprKind::Ident(name) if name.eq_ignore_ascii_case("Items"))
                && self
                    .current_class
                    .as_deref()
                    .and_then(|class_name| {
                        self.namespace_tree_instance_method_owner(class_name, "Items", 0)
                    })
                    .is_some_and(|owner| {
                        owner.to_ascii_lowercase().contains("observablecollection")
                    })
            {
                if self.emit_self_ref() {
                    self.emit_common("dotnet.observable_collection_items", 1, self.line);
                    self.compile_expr(&arg_exprs[0])?;
                    common::collections::emit_push(&mut self.chunks, self.current, self.line);
                    self.emit(Op::DROP);
                    self.emit(Op::NULL);
                    return Ok(());
                }
            }

            let class_name = if self.profile.namespaces.use_dotnet {
                resolve_receiver_type_hint(self, object)
            } else {
                None
            };
            if self.profile.namespaces.use_dotnet
                && field.eq_ignore_ascii_case("Reverse")
                && arg_exprs.len() == 2
                && !self.direct_receiver_has_own_pending_method(object, field)
            {
                self.compile_expr(object)?;
                for a in &arg_exprs {
                    self.compile_expr(a)?;
                }
                let line = self.line;
                self.emit_common("collections.reverse_range", 3, line);
                return Ok(());
            }
            if let Some(class_name) = class_name {
                if self
                    .resolve_pending_class_name_for_type_hint(&class_name)
                    .is_some()
                {
                    // User-defined classes win over shared .NET surface names
                    // like `Stack`, `Queue`, or `Dictionary`.
                } else {
                    let class_name = Self::normalize_type_hint(&class_name);
                    if let Some(target) = vybe_runtime::namespaces::lookup_type_instance_target(
                        &self.profile.namespaces.type_scopes,
                        &class_name,
                        field,
                        arg_exprs.len() as u8,
                    ) {
                        if self.profile.namespaces.use_dotnet && field.eq_ignore_ascii_case("Add") {
                            if let ExprKind::Index {
                                object: indexed_owner,
                                index,
                                ..
                            } = &object.kind
                            {
                                let owner_slot = self.define_local("__dotnet_index_owner");
                                let key_slot = self.define_local("__dotnet_index_key");
                                let value_slot = self.define_local("__dotnet_index_value");

                                self.compile_expr(indexed_owner)?;
                                self.emit_u16(Op::LOCAL_SET, owner_slot);
                                self.compile_collection_key(indexed_owner, index)?;
                                self.emit_u16(Op::LOCAL_SET, key_slot);

                                self.emit_u16(Op::LOCAL_GET, owner_slot);
                                self.emit_u16(Op::LOCAL_GET, key_slot);
                                common::collections::emit_get(
                                    &mut self.chunks,
                                    self.current,
                                    self.line,
                                );
                                self.emit_u16(Op::LOCAL_SET, value_slot);

                                self.emit_u16(Op::LOCAL_GET, value_slot);
                                for a in &arg_exprs {
                                    self.compile_expr(a)?;
                                }
                                let total_argc = (arg_exprs.len() + 1) as u8;
                                match &target {
                                    vybe_runtime::component_model::InstanceMethodTarget::Host {
                                        module,
                                        func,
                                        ..
                                    } => {
                                        let idx = self.import(module, func);
                                        self.emit_host_call(idx, total_argc);
                                    }
                                    vybe_runtime::component_model::InstanceMethodTarget::Common {
                                        emit, ..
                                    } => {
                                        let line = self.line;
                                        self.emit_common(emit, total_argc, line);
                                    }
                                }
                                self.emit(Op::DROP);

                                self.emit_u16(Op::LOCAL_GET, owner_slot);
                                self.emit_u16(Op::LOCAL_GET, key_slot);
                                self.emit_u16(Op::LOCAL_GET, value_slot);
                                common::collections::emit_set(
                                    &mut self.chunks,
                                    self.current,
                                    self.line,
                                );
                                self.emit(Op::DROP);
                                self.emit(Op::NULL);
                                return Ok(());
                            }
                        }

                        if matches!(&target, vybe_runtime::component_model::InstanceMethodTarget::Common { emit, .. } if emit == "collections.sort")
                            && arg_exprs.is_empty()
                        {
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
                            common::collections::emit_sort_with_comparator(
                                &mut self.chunks,
                                self.current,
                                self.line,
                            );
                            return Ok(());
                        }

                        if matches!(&target, vybe_runtime::component_model::InstanceMethodTarget::Common { emit, .. } if emit == "dotnet.array_sort")
                            && arg_exprs.len() == 1
                            && class_name.rsplit('.').next().is_some_and(|name| {
                                name.eq_ignore_ascii_case("List")
                                    || name.eq_ignore_ascii_case("ArrayList")
                            })
                            && matches!(&arg_exprs[0].kind, ExprKind::Lambda { params, .. } if params.len() == 2)
                        {
                            self.compile_expr(object)?;
                            self.compile_expr(&arg_exprs[0])?;
                            common::collections::emit_sort_with_comparator(
                                &mut self.chunks,
                                self.current,
                                self.line,
                            );
                            return Ok(());
                        }

                        self.compile_expr(object)?;
                        for a in &arg_exprs {
                            self.compile_expr(a)?;
                        }
                        let total_argc = (arg_exprs.len() + 1) as u8;
                        match target {
                            vybe_runtime::component_model::InstanceMethodTarget::Host {
                                module,
                                func,
                                ..
                            } => {
                                let idx = self.import(&module, &func);
                                self.emit_host_call(idx, total_argc);
                            }
                            vybe_runtime::component_model::InstanceMethodTarget::Common {
                                emit,
                                ..
                            } => {
                                let line = self.line;
                                let emit = if arg_exprs.len() == 1
                                    && matches!(&arg_exprs[0].kind, ExprKind::Lambda { params, .. } if params.len() == 2)
                                    && emit == "dotnet.linq_skip_while"
                                {
                                    "dotnet.linq_skip_while_indexed"
                                } else if arg_exprs.len() == 1
                                    && matches!(&arg_exprs[0].kind, ExprKind::Lambda { params, .. } if params.len() == 2)
                                    && emit == "dotnet.linq_take_while"
                                {
                                    "dotnet.linq_take_while_indexed"
                                } else if arg_exprs.len() == 1 && emit == "dotnet.linq_distinct" {
                                    "dotnet.linq_distinct_comparer"
                                } else if arg_exprs.len() == 2 && emit == "dotnet.linq_distinct_by"
                                {
                                    "dotnet.linq_distinct_by_comparer"
                                } else {
                                    emit.as_str()
                                };
                                self.emit_common(emit, total_argc, line);
                            }
                        }
                        return Ok(());
                    }
                }
            }
        }
        if let ExprKind::Member {
            object,
            field,
            null_safe,
        } = &callee.kind
        {
            if self.profile.ecma_promise_methods && !*null_safe {
                if let ExprKind::Ident(class_name) = &object.kind {
                    let class_canon = self.canon(class_name);
                    if (self.defined_classes.contains(&class_canon)
                        || self.pending_classes.contains_key(class_canon.as_str()))
                        && self.class_extends_builtin(&class_canon, "Promise")
                    {
                        let host_name = match field.as_str() {
                            "resolve" | "reject" | "all" | "race" | "allSettled" | "any"
                            | "try" => Some(field.as_str()),
                            _ => None,
                        };
                        if let Some(host_name) = host_name {
                            if let Some(arg) = arg_exprs.first() {
                                self.compile_expr(arg)?;
                            } else {
                                self.emit_const(Value::Undefined);
                            }
                            let idx = self.import("ecma:promise", host_name);
                            self.emit_host_call(idx, 1);
                            let promise_slot = self.define_local("__promise_subclass_result");
                            self.emit_u16(Op::LOCAL_SET, promise_slot);

                            self.emit_u16(Op::LOCAL_GET, promise_slot);
                            self.emit_var_get(&class_canon);
                            let proto_key = self.str_const("prototype");
                            self.emit_u16(Op::STRUCT_GET, proto_key);
                            let proto_link_key = self.str_const("__proto__");
                            self.emit_u16(Op::STRUCT_SET, proto_link_key);
                            self.emit(Op::DROP);

                            self.emit_u16(Op::LOCAL_GET, promise_slot);
                            return Ok(());
                        }
                    }
                }
            }
        }
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            let source_member_parts = self.flatten_member_chain(callee);
            if source_member_parts.len() >= 2 {
                if let Some(source_function) =
                    self.resolve_namespaced_function_identity(&source_member_parts.join("."))
                {
                    let global_idx = self.str_const(&source_function);
                    self.emit_u16(Op::GLOBAL_GET, global_idx);
                    for a in &arg_exprs {
                        self.compile_expr(a)?;
                    }
                    self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                    return Ok(());
                }
            }
            if resolves_to_static_container_method(self, object, field) {
                self.compile_expr(object)?;
                let obj_tmp = self.define_local("__static_container_obj");
                self.emit_u16(Op::LOCAL_SET, obj_tmp);
                let fn_tmp = self.define_local("__static_container_fn");
                let class_canon = self.canon(&self.flatten_member_chain(object).join("."));
                if self.profile.supports_private_fields && field.starts_with('#') {
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
                if self.profile.namespaces.use_dotnet && args.len() == 1 && !args[0].spread {
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
                if self.profile.has_ecma_globals && obj_name == "Object" && field == "fromEntries" {
                    if let Some(first) = arg_exprs.first() {
                        self.compile_expr(first)?;
                        if self.profile.has_generators {
                            common::collections::emit_spread_iterable(
                                &mut self.chunks,
                                self.current,
                                self.line,
                            );
                        }
                    } else {
                        let line = self.line;
                        common::expressions::emit_undefined(self.chunk(), line);
                    }
                    let idx = self.import("ecma:object", "fromEntries");
                    self.emit_host_call(idx, 1);
                    return Ok(());
                }
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
                // namespaceplan.md: migrated profiles gate this path through
                // the common resolver (adds locals-shadow correctness); the
                // direct map read serves the not-yet-migrated languages.
                let ns_module = if self.profile.uses_common_resolver {
                    match self.resolve_namespace_name(obj_name) {
                        Some(super::resolver::Resolution::NamespaceAlias { module }) => {
                            Some(module)
                        }
                        _ => None,
                    }
                } else {
                    self.host_namespace_aliases.get(&key).cloned()
                };
                if let Some(ns_module) = ns_module {
                    // §16.2: namespace member access is a COMPILE-TIME
                    // binding. Resolve the export statically — through the
                    // profile's mount-with-rename surface first
                    // (`j.dumps` → ecma:json/stringify), then the host's
                    // component-model export table — and emit a direct
                    // CALL_IMPORT. Only an export unknown at compile time
                    // falls back to the runtime namespace object.
                    let static_target = self
                        .module_exports
                        .get(&ns_module)
                        .and_then(|exports| exports.get(field))
                        .cloned()
                        .or_else(|| {
                            let ctx =
                                crate::primitives::instructions::host::CapabilityContext::get();
                            if ctx.functions.has(&ns_module, field) {
                                Some((ns_module.clone(), field.clone()))
                            } else {
                                None
                            }
                        });
                    if let Some((target_module, target_name)) = static_target {
                        for arg in &arg_exprs {
                            self.compile_expr(arg)?;
                        }
                        let idx = self.import(&target_module, &target_name);
                        self.emit_host_call(idx, arg_exprs.len() as u8);
                        return Ok(());
                    }
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
                // Component-model namespace fallback. Profile tree mounts
                // (`System` -> `dotnet.system`, `Flutter` -> `flutter`, ...)
                // drive this; a user binding on the leading ident shadows
                // namespace resolution in any language.
                let head_key = self.canon(&parts[0]);
                let mounted_tree_head = self.tree_mounts.contains_key(&head_key);
                if self.profile.uses_namespace_resolver()
                    && (!self.has_accessible_local_binding(&parts[0]) || mounted_tree_head)
                    && self.try_compile_namespace_component_call(&parts, &arg_exprs)?
                {
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
                    // namespaceplan.md migration: profiles that opted into
                    // the common resolver take the data-driven package-root
                    // path (profile `[[esm_default]] kind = "package-root"`
                    // mounts); the hardcoded arm below serves the languages
                    // that haven't migrated yet and dies with the last one.
                    let resolved = if self.profile.uses_common_resolver {
                        match self.resolve_namespace_path(&[prefix, inner_field, field]) {
                            Some(super::resolver::Resolution::HostImport { module, func }) => {
                                Some((module, func))
                            }
                            // The global namespace tree mounts EVERY host
                            // export (`vybe.gui.*` right next to `ecma.*`) —
                            // a tree HostCall is the same direct
                            // component-model call, no per-profile
                            // package-root data required.
                            Some(super::resolver::Resolution::Tree(
                                crate::primitives::namespaces::ResolutionTarget::HostCall {
                                    module,
                                    func,
                                    ..
                                },
                            )) => Some((module, func)),
                            _ => None,
                        }
                    } else {
                        let prefix_lc = self.canon(prefix);
                        if matches!(prefix_lc.as_str(), "vybe" | "wasi" | "wasm") {
                            Some((
                                format!("{}:{}", prefix_lc, self.canon(inner_field)),
                                field.clone(),
                            ))
                        } else {
                            None
                        }
                    };
                    if let Some((module, func)) = resolved {
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
                        let idx = self.import(&module, &func);
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
                    // A source type alias (`use App\Http\Request;`,
                    // VB `Imports X = …`) resolves the class path to its
                    // canonical (namespace-qualified) identity first, so
                    // `Str::upper()` on a use-aliased class dispatches to
                    // the declared `vendor.support.str`.
                    let aliased_path = self.resolve_source_type_alias(&class_path);
                    let full_canon = self.canon(&aliased_path);
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
                if self.profile.has_function_prototype_bind
                    && matches!(method_name.as_str(), "bind" | "call" | "apply")
                {
                    if let Some(class_canon) = early_static_class_canon.as_ref() {
                        let canon_field = self.canon(&method_name);
                        let shadowed =
                            self.pending_classes
                                .get(class_canon.as_str())
                                .is_some_and(|pc| {
                                    pc.static_method_overloads.contains_key(&canon_field)
                                });
                        if !shadowed {
                            early_static_class_canon = None;
                        }
                    }
                }

                if early_static_class_canon.is_some() && self.profile.uses_namespace_resolver() {
                    super::resolver::register_platform_trees();
                    let arity_tree_backed = vybe_runtime::namespaces::lookup_type_static_member(
                        &self.profile.namespaces.type_scopes,
                        &class_parts.join("."),
                        &method_name,
                    )
                    .and_then(|member| {
                        vybe_runtime::namespaces::select_overload(&member, arg_exprs.len() as u8)
                            .cloned()
                    })
                    .is_some();
                    let tree_backed = matches!(
                        self.resolve_profile_namespace_chain(&parts),
                        Some(super::resolver::Resolution::HostImport { .. })
                            | Some(super::resolver::Resolution::Tree(
                                crate::primitives::namespaces::ResolutionTarget::CommonEmit(_)
                            ))
                            | Some(super::resolver::Resolution::Tree(
                                crate::primitives::namespaces::ResolutionTarget::HostCall { .. }
                            ))
                            | Some(super::resolver::Resolution::Tree(
                                crate::primitives::namespaces::ResolutionTarget::Const(_)
                            ))
                            | Some(super::resolver::Resolution::ResolvedPrefix { .. })
                    ) || arity_tree_backed;
                    if tree_backed {
                        early_static_class_canon = None;
                    }
                }

                if let Some(class_canon) = early_static_class_canon {
                    let cls_idx = self.str_const(&class_canon);
                    self.emit_u16(Op::GLOBAL_GET, cls_idx);
                    let method_canon = self.canon(&method_name);
                    let qualified_method = self.canon(&format!("{}.{}", class_canon, method_name));
                    let method_idx = self.str_const(&method_canon);
                    self.emit_u16(Op::STRUCT_GET, method_idx);
                    // A FRESH slot per call site: reusing a same-named
                    // `__early_static_fn` slot aliases the outer callee when a
                    // nested call (`f(g(x))`) resolves the same name, so the
                    // outer `call_ref` would invoke the inner function. Each
                    // call must hold its own funcref until it fires.
                    let fn_tmp = self.define_local("__early_static_fn");
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

                    if self.profile.namespaces.use_dotnet && args.len() == 1 && !args[0].spread {
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
                            if self.class_prototype_dispatch() {
                                self.emit_js_receiver_host_or_bound_this_call(
                                    fn_tmp,
                                    receiver_slot,
                                    &arg_slots,
                                );
                            } else {
                                self.emit_call_ref_with_arg_slots(
                                    fn_tmp,
                                    Some(receiver_slot),
                                    &arg_slots,
                                );
                            }
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

                if let Some(source_function) =
                    self.resolve_namespaced_function_identity(&parts.join("."))
                {
                    let global_idx = self.str_const(&source_function);
                    self.emit_u16(Op::GLOBAL_GET, global_idx);
                    for a in &arg_exprs {
                        self.compile_expr(a)?;
                    }
                    self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                    return Ok(());
                }

                // Use the shared namespace resolver when profile data enables it.
                if self.profile.uses_namespace_resolver() {
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
                        // namespaceplan.md: platform surfaces are data in the
                        // shared tree; the common resolver handles the mounted chain.
                        let resolution = self.resolve_profile_namespace_chain(&parts);

                        match resolution {
                            Some(super::resolver::Resolution::GlobalAccess { name }) => {
                                let global_idx = self.str_const(&name);
                                self.emit_u16(Op::GLOBAL_GET, global_idx);
                                for a in &arg_exprs {
                                    self.compile_expr(a)?;
                                }
                                self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                                return Ok(());
                            }
                            Some(super::resolver::Resolution::Tree(
                                crate::primitives::namespaces::ResolutionTarget::CommonEmit(emit),
                            )) => {
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

                                if (emit.eq_ignore_ascii_case("dotnet.console_writeline")
                                    || emit.eq_ignore_ascii_case("dotnet.console_write"))
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
                            Some(
                                super::resolver::Resolution::HostImport { module, func }
                                | super::resolver::Resolution::Tree(
                                    crate::primitives::namespaces::ResolutionTarget::HostCall {
                                        module,
                                        func,
                                        ..
                                    },
                                ),
                            ) => {
                                if self.profile.namespaces.use_dotnet
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
                            Some(super::resolver::Resolution::NamespaceChain {
                                parts: ns_parts,
                            }) => {
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
                                    .map(|name| {
                                        vybe_runtime::namespaces::lookup_type_static_member(
                                            &self.profile.namespaces.type_scopes,
                                            name,
                                            name,
                                        )
                                        .is_some()
                                    })
                                    .unwrap_or(false);
                                if !is_const {
                                    for a in &arg_exprs {
                                        self.compile_expr(a)?;
                                    }
                                    self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                                }
                                return Ok(());
                            }
                            Some(super::resolver::Resolution::ScopedMember { local, members }) => {
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
                                            // Through the emit registry, like
                                            // every other platform emit — the
                                            // name is already registered by
                                            // the dotnet dispatch table.
                                            self.emit_common(
                                                "dotnet.process_wait_for_exit",
                                                1,
                                                line,
                                            );
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
                            Some(super::resolver::Resolution::NoOp) => {
                                self.emit(Op::NULL);
                                return Ok(());
                            }
                            _ => {
                                // Fall through to value methods and other resolution
                            }
                        }
                    }
                }

                // Non-dotnet: namespace aliases (JS: console → wasi:cli).
                // Reads from `host_namespace_aliases` (populated by the
                // Linker) instead of `profile.lookup_module_alias` — one
                // source of truth for Member-chain resolution.
                let mounted_tree_chain = self.profile.uses_namespace_resolver()
                    && self.resolve_profile_namespace_chain(&parts).is_some();
                if !mounted_tree_chain {
                    let alias_key = self.canon(&lower_parts[0]);
                    // namespaceplan.md: migrated profiles resolve the chain
                    // head through the common resolver (locals shadow);
                    // direct map read serves not-yet-migrated languages.
                    let ns_module = if self.profile.uses_common_resolver {
                        match self.resolve_namespace_name(&parts[0]) {
                            Some(super::resolver::Resolution::NamespaceAlias { module }) => {
                                Some(module)
                            }
                            _ => None,
                        }
                    } else {
                        self.host_namespace_aliases.get(&alias_key).cloned()
                    };
                    if let Some(module) = ns_module {
                        let is_prototype_chain = self.class_prototype_dispatch()
                            && lower_parts.len() > 2
                            && lower_parts
                                .get(1)
                                .is_some_and(|part| part.eq_ignore_ascii_case("prototype"));
                        let is_function_helper_chain = self.profile.has_function_prototype_bind
                            && lower_parts.len() > 2
                            && lower_parts.last().is_some_and(|part| {
                                matches!(part.as_str(), "call" | "apply" | "bind")
                            });
                        if is_prototype_chain || is_function_helper_chain {
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
                            // Mount-with-rename (namespaceplan.md): the
                            // profile's module-export surface reconciles the
                            // source-level name with the canonical host
                            // export (`json.dumps` / `j.dumps` →
                            // ecma:json/stringify) before emitting.
                            let (module, func) = match self
                                .module_exports
                                .get(&module)
                                .and_then(|exports| exports.get(&func))
                            {
                                Some((target_module, target_name)) => {
                                    (target_module.clone(), target_name.clone())
                                }
                                None => (module, func),
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
            let source_member_parts = self.flatten_member_chain(callee);
            if source_member_parts.len() >= 2 {
                if let Some(source_function) =
                    self.resolve_namespaced_function_identity(&source_member_parts.join("."))
                {
                    let global_idx = self.str_const(&source_function);
                    self.emit_u16(Op::GLOBAL_GET, global_idx);
                    for a in &arg_exprs {
                        self.compile_expr(a)?;
                    }
                    self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                    return Ok(());
                }
            }
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
                if self.profile.has_function_prototype_bind
                    && matches!(field.as_str(), "bind" | "call" | "apply")
                {
                    if let Some(canon) = static_class_canon.as_ref() {
                        let canon_field = self.canon(field);
                        let shadowed = self.pending_classes.get(canon.as_str()).is_some_and(|pc| {
                            pc.static_method_overloads.contains_key(&canon_field)
                        });
                        if !shadowed {
                            static_class_canon = None;
                        }
                    }
                }

                if let Some(canon) = static_class_canon {
                    if self.class_prototype_dispatch() {
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

                    if self.profile.namespaces.use_dotnet && args.len() == 1 && !args[0].spread {
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
            // Whether `.call`/`.apply` on a receiver means "invoke this
            // callable" is a LANGUAGE PROPERTY, declared in the profile.
            // Previously this was carved out as `profile.name == "java"`,
            // which both violated the no-language-names rule for shared code
            // and hid the general defect: every language that does NOT have
            // these members still routed them to the function builtins, so a
            // user method named `apply` returned null and one named `call`
            // panicked the host — in Dart, PHP, Python and JS alike.
            let invocation_members = self.profile.function_invocation_members;
            if self.profile.has_function_prototype_bind
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
            if invocation_members
                && !self.direct_receiver_has_own_pending_method(object, field)
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
            if self.profile.namespaces.use_dotnet
                && field.eq_ignore_ascii_case("Reverse")
                && arg_exprs.len() == 2
                && !self.direct_receiver_has_own_pending_method(object, field)
            {
                self.compile_expr(object)?;
                for a in &arg_exprs {
                    self.compile_expr(a)?;
                }
                let line = self.line;
                self.emit_common("collections.reverse_range", 3, line);
                return Ok(());
            }
            // Resolve the type to look up on the shared .NET surface. A typed
            // receiver uses its own type (unless it names a user class, which
            // wins over shared names like `Stack`/`Dictionary`). An *untyped*
            // receiver on the .NET path falls back to `IEnumerable`, so LINQ
            // resolves by method name on array literals, generator results,
            // and cross-language iterables — the adapter drains any iterable
            // at runtime via `generators.rs`. That single surface replaces the
            // per-function LINQ entries every .NET profile used to carry. The
            // fallback is skipped for user-shadowed names and for methods the
            // runtime collection registry owns at this arity (`List.Count()`).
            // Framework method resolution walks the .NET class hierarchy
            // (user subclasses into the descriptor) and wins over the "pending
            // class ⇒ dynamic" gate, so `Button.Show` / inherited control
            // members resolve to `vybe:gui` host calls instead of needing an
            // emitted thunk. Returns `None` on a user override → dynamic.
            let framework_method_owner = class_name
                .as_deref()
                .filter(|_| self.profile.namespaces.use_dotnet)
                .and_then(|cn| {
                    self.namespace_tree_instance_method_owner(cn, field, arg_exprs.len() as u8)
                });
            let surface_type: Option<String> = match &class_name {
                _ if framework_method_owner.is_some() => framework_method_owner,
                Some(cn) if self.resolve_pending_class_name_for_type_hint(cn).is_some() => None,
                // The .NET surface serves .NET-shaped profiles ONLY. Ungated,
                // this hijacked typed receivers in other languages (JS
                // `const a=[1,2]; a.reverse()` → the surface's non-mutating
                // Array.Reverse) — the exact "dotnet adapter leaked into
                // compiler core" disease namespaceplan.md documents.
                Some(cn) if !self.profile.namespaces.type_scopes.is_empty() => {
                    Some(Self::normalize_type_hint(cn))
                }
                Some(_) => None,
                None if self.profile.namespaces.use_dotnet
                    && !self.direct_receiver_has_own_pending_method(object, field)
                    && (is_dotnet_linq_method_name(field)
                        || !self.defined_class_methods.contains(&self.canon(field)))
                    && (is_dotnet_linq_method_name(field)
                        || !vybe_runtime::namespaces::scope_declares_member_arity(
                            &scope_segments(&self.profile.namespaces.runtime_collection_scope),
                            field,
                            arg_exprs.len() as u8,
                        )) =>
                {
                    Some("IEnumerable".to_string())
                }
                None => None,
            };
            if let Some(class_name) = surface_type {
                {
                    let owner = if class_name.eq_ignore_ascii_case("IEnumerable") {
                        self.namespace_tree_instance_method_owner(
                            &class_name,
                            field,
                            arg_exprs.len() as u8,
                        )
                        .unwrap_or_else(|| class_name.clone())
                    } else {
                        class_name.clone()
                    };
                    let prefer_linq_extension = is_dotnet_linq_method_name(field)
                        && (owner.eq_ignore_ascii_case("Array")
                            || owner.ends_with("[]")
                            || owner.ends_with("()"));
                    let target = if prefer_linq_extension {
                        vybe_runtime::namespaces::lookup_type_instance_target(
                            &self.profile.namespaces.type_scopes,
                            "IEnumerable",
                            field,
                            arg_exprs.len() as u8,
                        )
                    } else {
                        vybe_runtime::namespaces::lookup_type_instance_target(
                            &self.profile.namespaces.type_scopes,
                            &owner,
                            field,
                            arg_exprs.len() as u8,
                        )
                    }
                    .or_else(|| {
                        if is_dotnet_linq_method_name(field)
                            && !owner.eq_ignore_ascii_case("IEnumerable")
                        {
                            vybe_runtime::namespaces::lookup_type_instance_target(
                                &self.profile.namespaces.type_scopes,
                                "IEnumerable",
                                field,
                                arg_exprs.len() as u8,
                            )
                        } else {
                            None
                        }
                    });
                    if let Some(target) = target {
                        if matches!(&target, vybe_runtime::component_model::InstanceMethodTarget::Common { emit, .. } if emit == "collections.sort")
                            && arg_exprs.is_empty()
                        {
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
                            common::collections::emit_sort_with_comparator(
                                &mut self.chunks,
                                self.current,
                                self.line,
                            );
                            return Ok(());
                        }

                        if matches!(&target, vybe_runtime::component_model::InstanceMethodTarget::Common { emit, .. } if emit == "dotnet.array_sort")
                            && arg_exprs.len() == 1
                            && class_name.rsplit('.').next().is_some_and(|name| {
                                name.eq_ignore_ascii_case("List")
                                    || name.eq_ignore_ascii_case("ArrayList")
                            })
                            && matches!(&arg_exprs[0].kind, ExprKind::Lambda { params, .. } if params.len() == 2)
                        {
                            self.compile_expr(object)?;
                            self.compile_expr(&arg_exprs[0])?;
                            common::collections::emit_sort_with_comparator(
                                &mut self.chunks,
                                self.current,
                                self.line,
                            );
                            return Ok(());
                        }

                        // Compile receiver, then args.
                        self.compile_expr(object)?;
                        for a in &arg_exprs {
                            self.compile_expr(a)?;
                        }
                        let total_argc = (arg_exprs.len() + 1) as u8;
                        match target {
                            vybe_runtime::component_model::InstanceMethodTarget::Host {
                                module,
                                func,
                                ..
                            } => {
                                let idx = self.import(&module, &func);
                                self.emit_host_call(idx, total_argc);
                            }
                            vybe_runtime::component_model::InstanceMethodTarget::Common {
                                emit,
                                ..
                            } => {
                                let line = self.line;
                                let emit = if arg_exprs.len() == 1
                                    && matches!(&arg_exprs[0].kind, ExprKind::Lambda { params, .. } if params.len() == 2)
                                    && emit == "dotnet.linq_skip_while"
                                {
                                    "dotnet.linq_skip_while_indexed"
                                } else if arg_exprs.len() == 1
                                    && matches!(&arg_exprs[0].kind, ExprKind::Lambda { params, .. } if params.len() == 2)
                                    && emit == "dotnet.linq_take_while"
                                {
                                    "dotnet.linq_take_while_indexed"
                                } else if arg_exprs.len() == 1 && emit == "dotnet.linq_distinct" {
                                    "dotnet.linq_distinct_comparer"
                                } else if arg_exprs.len() == 2 && emit == "dotnet.linq_distinct_by"
                                {
                                    "dotnet.linq_distinct_by_comparer"
                                } else {
                                    emit.as_str()
                                };
                                self.emit_common(emit, total_argc, line);
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
                    // builtinslotplan.md step 5 — the table decides. Applies
                    // ONLY when the profile declared `slot = "..."` on this
                    // method AND the receiver's built-in type is statically
                    // known AND that pair is bound; otherwise the def is
                    // returned untouched, so a language that declares no slot
                    // cannot be affected.
                    .map(|def| self.apply_builtin_slot_binding(object, def))
            };
            // builtinslotplan.md step 3 — CENSUS, not a decision. Records which
            // `(built-in receiver type, method)` pairs actually reach
            // value-method dispatch, so steps 4-5 flip a measured list rather
            // than a guessed one. Emits nothing and changes nothing; off unless
            // VYBE_SLOT_AUDIT is set.
            if let Some(def) = matched_value_method.as_ref() {
                self.audit_builtin_slot_census(object, field, &def.emit);
            }
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
            // A receiver statically known to be a string keeps its string
            // value-method (`Contains` → str_includes, etc.). Names like
            // `Contains`/`IndexOf` collide with collection descriptor methods,
            // so they otherwise divert to runtime collection dispatch — which
            // does a dynamic method lookup on the receiver *object*. A string
            // is a primitive with no such method, so that path fails at
            // runtime ("undefined is not callable"). Strings are never runtime
            // collections, so the value-method is unambiguously correct here.
            let receiver_is_known_string = receiver_type_hint
                .as_deref()
                .map(Self::normalize_type_hint)
                .is_some_and(|type_hint| Self::is_string_type_hint(&type_hint));
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
            let java_direct_user_typed_receiver = self.profile.name == "java"
                && receiver_is_direct
                && receiver_is_user_type
                && field != "toString";
            // `defined_class_methods` is a FLAT, class-less set of every method
            // name declared by any class (`link.rs`). It predates the
            // declaration pass and exists only because `pending_classes` used
            // to be empty at call sites. Now that classes are registered before
            // any body compiles, the class-associated answer above is
            // authoritative — and consulting the flat set on top of it is
            // actively wrong: it claims a class has an inherited framework
            // member (`notifyListeners`, `findRenderObject`) merely because
            // some unrelated class declares that name, diverting the call away
            // from the framework adapter.
            //
            // So the flat set stays as a fallback ONLY where the receiver's
            // class cannot be resolved (untyped locals). Delete it outright
            // once receiver typing covers those — flexclassplan.md §3a.
            let receiver_class_known = self
                .infer_expr_type_hint(object)
                .as_deref()
                .and_then(|type_hint| self.resolve_pending_class_name_for_type_hint(type_hint))
                .is_some();
            let user_method_shadow = self.direct_receiver_has_own_pending_method(object, field)
                || receiver_has_pending_user_method
                || java_direct_user_typed_receiver
                || (receiver_is_direct
                    && !receiver_class_known
                    && !receiver_is_known_builtin_value
                    && self.defined_class_methods.contains(&canon_field))
                || (receiver_is_direct
                    && !receiver_class_known
                    && receiver_is_user_type
                    && self.defined_class_methods.contains(&canon_field));
            let java_member_apply =
                self.profile.name == "java" && receiver_is_direct && field == "apply";
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
            if user_method_shadow || java_member_apply || is_array_method {
                // Fall through — let the HOF dispatch or generic call path handle it
            } else if array_only_value_method_for_non_array {
                // Array-only value methods like `.entries()` must not steal
                // Map/Set receivers away from runtime method dispatch.
            } else if self.profile.namespaces.use_dotnet
                && vybe_runtime::namespaces::scope_declares_member_arity(
                    &scope_segments(&self.profile.namespaces.runtime_collection_scope),
                    field,
                    arg_exprs.len() as u8,
                )
                && !prefer_dotnet_adapter
                && !(receiver_is_known_string && matched_value_method.is_some())
            {
                // Let the generic member-call path consult the runtime type
                // registry for shared .NET collection methods instead of
                // intercepting them via language profile value-method tables.
                // Exception: a statically-known string receiver keeps its
                // string value-method (strings are never runtime collections).
            } else if let Some(def) = matched_value_method {
                if self.profile.supports_spread_arguments
                    && field == "push"
                    && args.iter().any(|arg| arg.spread)
                {
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
                        crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
                    };
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    crate::primitives::ops::emit_dyn_not(self.chunk(), line);
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
                        crate::primitives::ops::emit_dyn_add(self.chunk(), line);
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
                // Intrinsic value methods (`s.capitalize()`, `s.removeprefix(p)`, …)
                // recompile from the AST with the receiver as `args[0]`; they must
                // NOT be handed a pre-pushed receiver/args like the opcode/common
                // path below (which would leave the receiver untouched on the stack).
                if let BuiltinEmit::Intrinsic(intrinsic_name) = &def.emit {
                    let name = intrinsic_name.clone();
                    let mut intr_args: Vec<&Expression> = Vec::with_capacity(arg_exprs.len() + 1);
                    intr_args.push(object);
                    intr_args.extend(arg_exprs.iter().copied());
                    self.emit_intrinsic(&name, &intr_args)?;
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
                        // "from start to end of string". wasm:js-string.substring
                        // wants `[s, start, end]`; default end to a
                        // sentinel large value (VM clamps to s.len()).
                        // Same shape applies to ECMA-262 §22.1.3.16
                        // `String.prototype.slice(start)`.
                        "strings.substring" | "strings.slice" if arg_exprs.len() < 2 => {
                            self.emit_const(Value::I32(i32::MAX));
                        }
                        // C#'s `string.ToCharArray()` lowers to a split which
                        // needs a delimiter on the stack. The .NET semantics
                        // ("each char one element") match splitting on the
                        // empty string.
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
            let requires_dynamic_callback_dispatch = self.profile.ecma_array_method_dispatch
                && arg_exprs
                    .first()
                    .is_some_and(|expr| matches!(expr.kind, ExprKind::Call { .. }));
            if !user_class_method
                && !receiver_is_url_search_params
                && !requires_dynamic_callback_dispatch
                && self.profile.lookup_array_method(&field_lower).is_some()
            {
                // (re-fetch only when we're committed to the HOF path so
                // the method name lookup matches the previous behaviour)
            }
            if let Some(stdlib_name) = self
                .profile
                .lookup_array_method(&field_lower)
                .filter(|_| {
                    !self.profile.ecma_array_method_dispatch
                        && !user_class_method
                        && !receiver_is_url_search_params
                        && !requires_dynamic_callback_dispatch
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
                                crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
                            };
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                            crate::primitives::ops::emit_dyn_not(self.chunk(), line);
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
                                crate::primitives::ops::emit_dyn_add(self.chunk(), line);
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
                        if self.profile.ecma_array_method_dispatch {
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
                        if self.profile.has_undefined_value {
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
                            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
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
                            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
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
                        if self.profile.ecma_array_method_dispatch {
                            // ecma:array.sort returns the sorted array
                            // (in-place, returns receiver). One-arg call.
                            let idx = self.import("ecma:array", "sort");
                            self.emit_u16(Op::LOCAL_GET, arr_slot);
                            self.emit_host_call(idx, 1);
                        } else {
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
                            common::collections::emit_sort_with_comparator(
                                &mut self.chunks,
                                self.current,
                                line,
                            );
                        }
                        self.chunk().emit_else(line);
                        // JS is 1-to-1 with the ECMA runtime: `Array.prototype.sort`
                        // (§23.1.3.30) IS `ecma:array/sort` — call it directly with
                        // the user comparator, no stdlib comparator polyfill.
                        let idx = self.import("ecma:array", "sort");
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_host_call(idx, 2);
                        self.chunk().emit_end(line);
                    }
                    "sort_by_key" => {
                        // .NET OrderBy(keySelector) — 1-arg key extractor
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit(Op::REF_IS_NULL);
                        let line = self.line;
                        self.chunk().emit_if_value(line);
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        common::collections::emit_sort(&mut self.chunks, self.current, line);
                        self.chunk().emit_else(line);
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        common::collections::emit_sort_by_key_in_place(
                            &mut self.chunks,
                            self.current,
                            line,
                        );
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
                            crate::primitives::ops::emit_dyn_ge(self.chunk(), line);
                        };
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        crate::primitives::ops::emit_dyn_not(self.chunk(), line);
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
                            crate::primitives::ops::emit_dyn_ge(self.chunk(), line);
                        };
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        crate::primitives::ops::emit_dyn_not(self.chunk(), line);
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
                            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
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
                            crate::primitives::ops::emit_dyn_ge(self.chunk(), line);
                        };
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        crate::primitives::ops::emit_dyn_not(self.chunk(), line);
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
                            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
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
                            crate::primitives::ops::emit_dyn_ge(self.chunk(), line);
                        };
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        crate::primitives::ops::emit_dyn_not(self.chunk(), line);
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
                            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
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
                            crate::primitives::ops::emit_dyn_add(self.chunk(), line);
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
            if self.private_member_access_forbidden(field) {
                self.emit_private_access_denied(field)?;
                return Ok(());
            }
            if self.profile.supports_private_fields && field.starts_with('#') && !*null_safe {
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

                let private_storage_name;
                if let Some(class_name) = static_class_canon {
                    if let Some(overload) =
                        self.resolve_static_method_overload_for_type(&class_name, field, &arg_exprs)
                    {
                        private_storage_name =
                            self.js_member_storage_name_for_class(&class_name, field);
                        self.emit_js_private_brand_check(obj_tmp, &private_storage_name)?;
                        let line = self.line;
                        self.emit_u16(Op::REF_FUNC, overload.chunk_idx as u16);
                        self.chunk().emit(0, line);
                        self.emit_u16(Op::LOCAL_SET, fn_tmp);
                    } else {
                        let field_name = self.js_member_storage_name_for_class(&class_name, field);
                        private_storage_name = field_name.clone();
                        self.emit_js_private_brand_check(obj_tmp, &private_storage_name)?;
                        let prop = self.str_const(&field_name);
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.emit_u16(Op::STRUCT_GET, prop);
                        self.emit_u16(Op::LOCAL_SET, fn_tmp);
                    }
                } else {
                    let field_name = self.js_member_storage_name_for_receiver(object, field);
                    private_storage_name = field_name.clone();
                    self.emit_js_private_brand_check(obj_tmp, &private_storage_name)?;
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
            if self.profile.ecma_promise_methods {
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
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), gen_if_line);
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
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), gen_if_line);
                    self.chunk().emit_if(gen_if_line);

                    let value_slot = self.define_local("__gen_return_value");
                    let done_slot = self.define_local("__gen_return_done");
                    let returned_key = self.str_const("__vybe_gen_returned");

                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    let is_done_idx = self.import("ecma:value", "isGeneratorDone");
                    self.emit_host_call(is_done_idx, 1);
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
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
                    crate::primitives::generators::emit_resume(self.chunk(), line);
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
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), gen_if_line);
                    // ASYNC generators skip this raw fast path — their
                    // attached `__vybe_async_generator_next` driver returns
                    // the §27.6.1.2 promise-wrapped IteratorResult (and
                    // rejects on a body throw); regular method dispatch
                    // below calls it.
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    let async_gen_key = self.str_const("__vybe_async_gen");
                    self.emit_u16(Op::STRUCT_GET, async_gen_key);
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), gen_if_line);
                    self.emit(Op::I32_EQZ);
                    self.emit(Op::I32_AND);
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
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
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
                        crate::primitives::generators::emit_next(self.chunk(), line);
                        let has_more_slot = self.define_local("__gen_has_more");
                        self.emit_u16(Op::LOCAL_SET, has_more_slot);
                        self.emit_u16(Op::LOCAL_SET, value_slot);
                        self.emit_u16(Op::LOCAL_GET, has_more_slot);
                        {
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                        };
                        {
                            let line = self.line;
                            // emit_dyn_not: has_more → i32 (1 if done, 0 if not done)
                            // emit_i32_to_bool: convert to Bool for ECMA `done` property
                            crate::primitives::ops::emit_dyn_not(self.chunk(), line);
                            crate::primitives::ops::emit_i32_to_bool(self.chunk(), line);
                        };
                        self.emit_u16(Op::LOCAL_SET, done_slot);
                        // Per ECMA-262 §27.5.3.5: when a generator completes
                        // (done=true) with no explicit return value, the VM
                        // leaves null on the stack. Convert null → undefined
                        // so the {value} field is spec-correct.
                        if self.profile.ecma_iterator_result_shape {
                            self.emit_u16(Op::LOCAL_GET, done_slot);
                            let line = self.line;
                            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
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
                        crate::primitives::generators::emit_resume(self.chunk(), line);
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
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), gen_if_line);
                    self.chunk().emit_if(gen_if_line);

                    let value_slot = self.define_local("__gen_throw_value");
                    let done_slot = self.define_local("__gen_throw_done");
                    let started_key = self.str_const("__vybe_gen_started");

                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit_u16(Op::STRUCT_GET, started_key);
                    {
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.emit(Op::I32_EQZ);
                    let line = self.line;
                    self.chunk().emit_if(line);

                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    let line = self.line;
                    crate::primitives::generators::emit_next(self.chunk(), line);
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
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    };
                    self.emit(Op::I32_EQZ);
                    let line = self.line;
                    self.chunk().emit_if(line);
                    if arg_exprs.is_empty() {
                        inst!(self, core_wasm::undefined);
                    } else {
                        self.compile_expr(&arg_exprs[0])?;
                    }
                    {
                        let line = self.line;
                        crate::primitives::errors::emit_throw(self.chunk(), line);
                    }
                    self.chunk().emit_end(line);
                    self.chunk().emit_end(line);

                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    if arg_exprs.is_empty() {
                        inst!(self, core_wasm::undefined);
                    } else {
                        self.compile_expr(&arg_exprs[0])?;
                    }
                    let line = self.line;
                    crate::primitives::generators::emit_resume_throw(self.chunk(), line);
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

                        // The member may not be there. Knowing the receiver's
                        // class does NOT mean knowing all its members: a class
                        // whose ancestry reaches a type the compiler doesn't
                        // model (a framework/catalog parent such as Flutter's
                        // `ChangeNotifier`) has a PARTIAL member list, and an
                        // inherited member is absent from it. Reading
                        // `fn.__vybe_method_receiver` before establishing that
                        // `fn` exists traps on undefined, so the lenient
                        // fallback below was unreachable exactly when it was
                        // needed. Guard the read.
                        let miss_line = self.line;
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit(Op::REF_IS_NULL);
                        self.chunk().emit_if_value(miss_line);
                        self.emit_js_lookup_or_invoke_method_call(
                            obj_tmp,
                            &method_name,
                            &arg_slots,
                        )?;
                        self.chunk().emit_else(miss_line);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        fn_call!(self, "wasm:js-undefined", "test", 1);
                        self.chunk().emit_if_value(miss_line);
                        self.emit_js_lookup_or_invoke_method_call(
                            obj_tmp,
                            &method_name,
                            &arg_slots,
                        )?;
                        self.chunk().emit_else(miss_line);

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
                        )?;
                        self.chunk().emit_else(line);
                        self.emit_u16(Op::LOCAL_GET, marker_slot);
                        fn_call!(self, "wasm:js-undefined", "test", 1);
                        let line = self.line;
                        self.chunk().emit_if_value(line);
                        self.emit_js_lookup_or_invoke_method_call(
                            obj_tmp,
                            &method_name,
                            &arg_slots,
                        )?;
                        self.chunk().emit_else(line);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        for slot in &arg_slots {
                            self.emit_u16(Op::LOCAL_GET, *slot);
                        }
                        self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                        self.chunk().emit_end(line);
                        self.chunk().emit_end(line);
                        // close the two member-present guards
                        self.chunk().emit_end(miss_line);
                        self.chunk().emit_end(miss_line);
                    } else {
                        self.emit_js_lookup_or_invoke_method_call(
                            obj_tmp,
                            &method_name,
                            &arg_slots,
                        )?;
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

                let receiver_is_pending_class = self
                    .infer_expr_type_hint(object)
                    .as_deref()
                    .and_then(|type_hint| self.resolve_pending_class_name_for_type_hint(type_hint))
                    .is_some();
                if receiver_is_pending_class {
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit_u16(Op::STRUCT_GET, prop);
                    let class_fn_slot = self.define_local("__js_class_dispatch_fn");
                    self.emit_u16(Op::LOCAL_SET, class_fn_slot);
                    if args.iter().any(|arg| arg.spread) {
                        let (args_slot, known_len) =
                            self.compile_call_args_array(args, "js_class_dispatch_spread")?;
                        self.emit_call_ref_with_args_array(
                            class_fn_slot,
                            Some(obj_tmp),
                            args_slot,
                            known_len,
                        );
                    } else {
                        let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                        for (index, arg) in arg_exprs.iter().enumerate() {
                            self.compile_expr(arg)?;
                            let arg_slot =
                                self.define_local(&format!("__js_class_dispatch_arg_{}", index));
                            self.emit_u16(Op::LOCAL_SET, arg_slot);
                            arg_slots.push(arg_slot);
                        }
                        // Prototype-dispatch profiles (ECMA-262 §15.7) declare no
                        // explicit receiver param — `this` arrives via the binding.
                        // Passing the receiver positionally would land it in
                        // argument 0 and shift every real argument
                        // (`c.f(7)` → f receives the object, not 7).
                        if self.class_prototype_dispatch() {
                            // Knowing the receiver's class does NOT mean the
                            // member is present on it. A class whose ancestry
                            // reaches a type the compiler doesn't model (a
                            // framework/catalog parent such as Flutter's
                            // `ChangeNotifier`) has a PARTIAL member list, so an
                            // INHERITED member reads as undefined here. Handing
                            // that to the receiver-marker dispatch calls
                            // `hasOwn(undefined, …)`, which throws.
                            //
                            // Guard the miss and fall back to dynamic lookup —
                            // the same shape the sibling dispatch site already
                            // uses. Before the declaration pass this was
                            // unreachable, because a call in `main` never
                            // resolved its receiver's class at all.
                            let miss_line = self.line;
                            self.emit_u16(Op::LOCAL_GET, class_fn_slot);
                            self.emit(Op::REF_IS_NULL);
                            self.chunk().emit_if(miss_line);
                            self.emit_js_lookup_or_invoke_method_call(obj_tmp, field, &arg_slots)?;
                            self.chunk().emit_else(miss_line);
                            self.emit_u16(Op::LOCAL_GET, class_fn_slot);
                            fn_call!(self, "wasm:js-undefined", "test", 1);
                            self.chunk().emit_if(miss_line);
                            self.emit_js_lookup_or_invoke_method_call(obj_tmp, field, &arg_slots)?;
                            self.chunk().emit_else(miss_line);
                            self.emit_js_receiver_host_or_bound_this_call(
                                class_fn_slot,
                                obj_tmp,
                                &arg_slots,
                            );
                            self.chunk().emit_end(miss_line);
                            self.chunk().emit_end(miss_line);
                        } else {
                            self.emit_call_ref_with_arg_slots(
                                class_fn_slot,
                                Some(obj_tmp),
                                &arg_slots,
                            );
                        }
                    }
                    self.emit_u16(Op::LOCAL_SET, js_result_slot);
                } else {
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
                }
                self.chunk().emit_end(line);
                self.emit_u16(Op::LOCAL_GET, js_result_slot);
                return Ok(());
            }

            self.compile_expr(object)?;
            let obj_tmp = self.define_local("__obj");
            self.reserve_local_slot(obj_tmp);
            self.emit_u16(Op::LOCAL_SET, obj_tmp);

            if self.profile.supports_private_fields && field.starts_with('#') && !*null_safe {
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
                    // C# delegate null-conditional invocation: `d?.Invoke(args)`.
                    // A multicast delegate is an array of handlers, so route
                    // through the shared invoker (iterates + calls each in
                    // order) rather than a bare CALL_REF that only handles a
                    // single function.
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    for a in &arg_exprs {
                        self.compile_expr(a)?;
                    }
                    common::delegates::emit_invoke(
                        &mut self.chunks,
                        self.current,
                        (arg_exprs.len() + 1) as u8,
                        self.line,
                    );
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

            if field.eq_ignore_ascii_case("Invoke") {
                let receiver_type_hint = self.infer_expr_type_hint(object);
                let receiver_is_delegate = receiver_type_hint
                    .as_deref()
                    .is_some_and(|type_hint| Self::is_callable_type_hint(type_hint))
                    || (self.profile.namespaces.use_dotnet
                        && receiver_type_hint.as_deref().is_some_and(|type_hint| {
                            let normalized = Self::normalize_type_hint(type_hint);
                            !self.defined_classes.contains(&self.canon(&normalized))
                                && !matches!(
                                    normalized.to_ascii_lowercase().as_str(),
                                    "object"
                                        | "system.object"
                                        | "string"
                                        | "system.string"
                                        | "integer"
                                        | "int"
                                        | "int32"
                                        | "system.int32"
                                        | "boolean"
                                        | "bool"
                                        | "system.boolean"
                                )
                        }))
                    || matches!(
                        object.kind,
                        ExprKind::Lambda { .. } | ExprKind::AddressOf(_)
                    );
                if receiver_is_delegate {
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    for arg in &arg_exprs {
                        self.compile_expr(arg)?;
                    }
                    common::delegates::emit_invoke(
                        &mut self.chunks,
                        self.current,
                        (arg_exprs.len() + 1) as u8,
                        self.line,
                    );
                    return Ok(());
                }
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
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                    self.chunk().emit_if(line);

                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    match field_name.as_str() {
                        "send" => {
                            self.compile_expr(&arg_exprs[0])?;
                        }
                        "throw" => {
                            self.compile_expr(&arg_exprs[0])?;
                            let line = self.line;
                            crate::primitives::generators::emit_resume_throw(self.chunk(), line);
                        }
                        "close" => {
                            self.emit(Op::NULL);
                            self.emit_generator_control_packet_from_stack("return");
                            let line = self.line;
                            crate::primitives::generators::emit_resume(self.chunk(), line);
                        }
                        _ => unreachable!(),
                    }
                    if field_name == "send" {
                        let line = self.line;
                        crate::primitives::generators::emit_resume(self.chunk(), line);
                    }
                    self.chunk().emit_end(line);
                }
            }

            if let Some(result_slot) = buffered_generator_end {
                if self.profile.namespaces.use_dotnet
                    && arg_exprs.is_empty()
                    && field.eq_ignore_ascii_case("sort")
                    && vybe_runtime::namespaces::scope_declares_member_arity(
                        &scope_segments(&self.profile.namespaces.runtime_collection_scope),
                        field,
                        0,
                    )
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
                if self.class_prototype_dispatch() {
                    let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                    for (index, arg) in arg_exprs.iter().enumerate() {
                        self.compile_expr(arg)?;
                        let arg_slot =
                            self.define_local(&format!("__js_member_bound_arg_{}", index));
                        self.emit_u16(Op::LOCAL_SET, arg_slot);
                        arg_slots.push(arg_slot);
                    }
                    self.emit_js_lookup_or_invoke_method_call(obj_tmp, field, &arg_slots)?;
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
                    if self.profile.namespaces.use_dotnet && args.len() == 1 && !args[0].spread {
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
                    } else if self.class_prototype_dispatch() {
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
                            crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                        };
                        let line = self.line;
                        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
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
                        crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                    };
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
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
                if self.class_prototype_dispatch() {
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
                        self.emit_js_lookup_or_invoke_method_call(obj_tmp, field, &arg_slots)?;
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
                            if args.iter().any(|arg| arg.by_ref) {
                                let pack_slot = self.define_local("__member_fast_by_ref_pack");
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
                                common::collections::emit_get(
                                    &mut self.chunks,
                                    self.current,
                                    self.line,
                                );
                            }
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
                        if self.class_prototype_dispatch() {
                            self.emit_js_receiver_host_or_bound_this_call(
                                fn_tmp,
                                receiver_slot,
                                &arg_slots,
                            );
                        } else {
                            self.emit_call_ref_with_arg_slots(
                                fn_tmp,
                                Some(receiver_slot),
                                &arg_slots,
                            );
                        }
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
                && vybe_runtime::namespaces::scope_declares_member_arity(
                    &scope_segments(&self.profile.namespaces.runtime_collection_scope),
                    field,
                    0,
                )
            {
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                let line = self.line;
                self.emit_common("dotnet.array_sort", 1, line);
                return Ok(());
            }

            if let Some(overload) =
                self.resolve_instance_method_overload(object, field, &arg_exprs, false)
            {
                let chunk_idx = overload.chunk_idx;
                if overload.is_virtual {
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit_autoderef_pointer_cell();
                    let overload_field =
                        self.overload_storage_name(&field_name, &overload.param_types);
                    let overload_prop = self.str_const(&overload_field);
                    self.emit_u16(Op::STRUCT_GET, overload_prop);
                    let virtual_fn_tmp = self.define_local("__virtual_instance_method_fn");
                    self.emit_u16(Op::LOCAL_SET, virtual_fn_tmp);
                    if overload.signature.has_rest {
                        self.emit_known_rest_call_from_local(
                            virtual_fn_tmp,
                            if self.class_prototype_dispatch() {
                                None
                            } else {
                                Some(obj_tmp)
                            },
                            args,
                            &overload.signature,
                        )?;
                    } else {
                        self.emit_instance_method_call_from_fn_slot(
                            virtual_fn_tmp,
                            field,
                            obj_tmp,
                            args,
                            &arg_exprs,
                        )?;
                    }
                    return Ok(());
                }
                if overload.signature.has_rest {
                    let line = self.line;
                    self.emit_u16(Op::REF_FUNC, chunk_idx as u16);
                    self.chunk().emit(0, line);
                    let direct_fn_tmp = self.define_local("__direct_instance_fn");
                    self.emit_u16(Op::LOCAL_SET, direct_fn_tmp);
                    self.emit_known_rest_call_from_local(
                        direct_fn_tmp,
                        if self.class_prototype_dispatch() {
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
                    crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                };
                let line = self.line;
                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
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
                if self.profile.namespaces.use_dotnet && args.len() == 1 && !args[0].spread {
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
                        crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                    };
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
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
                    crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                };
                let line = self.line;
                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
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
                    if self.class_prototype_dispatch() {
                        None
                    } else {
                        Some(obj_tmp)
                    },
                    args,
                    &overload.signature,
                )?;
            } else {
                let receiver_slot = obj_tmp;
                if self.class_prototype_dispatch() {
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
                                        || self
                                            .resolve_pending_class_name_for_type_hint(type_hint)
                                            .is_some()
                                },
                            );
                    if js_user_defined_member {
                        if args.iter().any(|arg| arg.spread) {
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
                            let (args_slot, known_len) =
                                self.compile_call_args_array(args, "js_member_call_spread")?;
                            self.emit_call_ref_with_args_array(
                                fn_tmp,
                                Some(obj_tmp),
                                args_slot,
                                known_len,
                            );
                            self.emit_u16(Op::LOCAL_SET, js_result_slot);
                            self.emit_const(Value::I32(1));
                            self.emit_u16(Op::LOCAL_SET, js_handled_slot);
                            self.chunk().emit_end(line);
                            self.chunk().emit_end(line);
                        } else {
                            let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                            for (index, arg) in arg_exprs.iter().enumerate() {
                                self.compile_expr(arg)?;
                                let arg_slot =
                                    self.define_local(&format!("__js_member_call_arg_{}", index));
                                self.emit_u16(Op::LOCAL_SET, arg_slot);
                                arg_slots.push(arg_slot);
                            }
                            self.emit_u16(Op::LOCAL_GET, fn_tmp);
                            self.emit(Op::REF_IS_NULL);
                            let line = self.line;
                            self.chunk().emit_if(line);
                            self.emit_js_lookup_or_invoke_method_call(obj_tmp, field, &arg_slots)?;
                            self.chunk().emit_else(line);
                            self.emit_u16(Op::LOCAL_GET, fn_tmp);
                            fn_call!(self, "wasm:js-undefined", "test", 1);
                            let line = self.line;
                            self.chunk().emit_if(line);
                            self.emit_js_lookup_or_invoke_method_call(obj_tmp, field, &arg_slots)?;
                            self.chunk().emit_else(line);
                            self.emit_js_receiver_host_or_bound_this_call(
                                fn_tmp, obj_tmp, &arg_slots,
                            );
                            self.chunk().emit_end(line);
                            self.chunk().emit_end(line);
                            self.emit_u16(Op::LOCAL_SET, js_result_slot);
                            self.emit_const(Value::I32(1));
                            self.emit_u16(Op::LOCAL_SET, js_handled_slot);
                        }
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
            if self.profile.has_function_constructor && name == "Function" {
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

            let resolved_call_name = self.resolve_namespaced_function_identity(name);
            let name = resolved_call_name.as_deref().unwrap_or(name.as_str());

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
            if let Some(target) = self.namespace_import_bindings.get(&key).cloned() {
                match target {
                    crate::primitives::namespaces::ResolutionTarget::CommonEmit(emit) => {
                        for a in &arg_exprs {
                            self.compile_expr(a)?;
                        }
                        let line = self.line;
                        self.emit_common(&emit, arg_exprs.len() as u8, line);
                        return Ok(());
                    }
                    crate::primitives::namespaces::ResolutionTarget::HostCall {
                        module,
                        func,
                        ..
                    } => {
                        for a in &arg_exprs {
                            self.compile_expr(a)?;
                        }
                        let idx = self.import(&module, &func);
                        self.emit_host_call(idx, arg_exprs.len() as u8);
                        return Ok(());
                    }
                    crate::primitives::namespaces::ResolutionTarget::Ctor {
                        spec: Some(spec),
                        ..
                    } => {
                        return self.emit_tree_ctor_construction(&spec, args);
                    }
                    _ => {}
                }
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
                        crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                    };
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
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

                        if self.profile.namespaces.use_dotnet && args.len() == 1 && !args[0].spread
                        {
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

            if !is_known_func {
                let is_delegate_typed = self.lookup_var_type_hint(name).is_some_and(|type_hint| {
                    Self::is_callable_type_hint(type_hint)
                        || (self.profile.namespaces.use_dotnet && {
                            let normalized = Self::normalize_type_hint(type_hint);
                            !self.defined_classes.contains(&self.canon(&normalized))
                                && !matches!(
                                    normalized.to_ascii_lowercase().as_str(),
                                    "object"
                                        | "system.object"
                                        | "string"
                                        | "system.string"
                                        | "integer"
                                        | "int"
                                        | "int32"
                                        | "system.int32"
                                        | "boolean"
                                        | "bool"
                                        | "system.boolean"
                                )
                        })
                });
                if is_delegate_typed {
                    self.emit_var_get(name);
                    for arg in &arg_exprs {
                        self.compile_expr(arg)?;
                    }
                    common::delegates::emit_invoke(
                        &mut self.chunks,
                        self.current,
                        (arg_exprs.len() + 1) as u8,
                        self.line,
                    );
                    return Ok(());
                }
            }

            // VB array access: `arr(idx)` when `arr` is a known data variable
            // (local OR top-level global from `Dim arr(5)`) and is NOT a
            // declared function or class. VB syntactically overloads `()` for
            // both calls and indexing — the disambiguator is whether the head
            // is a callable function or a value. We must exclude both
            // `defined_functions` and `defined_classes` from the "looks like
            // a variable" set, otherwise `GetResult()` (function call) and
            // `New Result()` (class) would be mis-identified as indexing.
            // `parens_for_index` alone is the gate — PHP never sets it (it
            // defaults false), so the PHP name check that used to sit here was
            // already unreachable.
            if !is_known_func && arg_exprs.len() == 1 && self.profile.parens_for_index {
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
                // A bare `Foo(...)` naming a declared class constructs it —
                // implicit-self applies to a class's own members, and a class
                // is not one of them. Without this, `Vec2(x + other.x)` inside
                // `Vec2 operator +(...)` compiles to `this.Vec2(...)`, whose
                // STRUCT_GET yields undefined at runtime. Static members
                // already escape via `current_member_is_static`.
                let is_known_class = self.defined_classes.contains(&self.canon(name));
                if !is_local && !is_known_func && !is_known_class {
                    if self.profile.namespaces.use_dotnet {
                        if let Some(current_class) = self.current_class.clone() {
                            if let Some(owner) = self.namespace_tree_instance_method_owner(
                                &current_class,
                                name,
                                arg_exprs.len() as u8,
                            ) {
                                let target = vybe_runtime::namespaces::lookup_type_instance_target(
                                    &self.profile.namespaces.type_scopes,
                                    &owner,
                                    name,
                                    arg_exprs.len() as u8,
                                );
                                if let Some(target) = target {
                                    if self.emit_self_ref() {
                                        for arg in &arg_exprs {
                                            self.compile_expr(arg)?;
                                        }
                                        let total_argc = (arg_exprs.len() + 1) as u8;
                                        match target {
                                            vybe_runtime::component_model::InstanceMethodTarget::Host {
                                                module,
                                                func,
                                                ..
                                            } => {
                                                let idx = self.import(&module, &func);
                                                self.emit_host_call(idx, total_argc);
                                            }
                                            vybe_runtime::component_model::InstanceMethodTarget::Common {
                                                emit,
                                                ..
                                            } => {
                                                let line = self.line;
                                                self.emit_common(&emit, total_argc, line);
                                            }
                                        }
                                        return Ok(());
                                    }
                                }
                            }
                        }
                    }
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
                self.emit_source_function_callable_name_resolution(callee_slot);
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
            // An object used as a callee: probe the Call SLOT before giving up.
            // Kotlin `operator fun invoke`, Python `__call__`, PHP `__invoke`,
            // Dart `call` and a C# `()` operator all fill it, so ONE probe
            // reaches every one of them — which is exactly why this must not be
            // gated on a language NAME. It was `is_python_profile()`, so
            // `val f = Box(); f("x")` trapped with "Not a function" in every
            // other language that has the feature.
            if self.profile.callable_objects && !is_known_func {
                let callee_slot = self.define_local("__py_call_target");
                self.emit_var_get(name);
                self.emit_u16(Op::LOCAL_SET, callee_slot);

                self.emit_u16(Op::LOCAL_GET, callee_slot);
                let typeof_idx = self.import("ecma:value", "typeof");
                self.emit_host_call(typeof_idx, 1);
                self.emit_const(Value::String(Arc::from("function")));
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
                };
                let line = self.line;
                crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
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
                // The Call SLOT. Python spells it `__call__`, PHP `__invoke`,
                // Dart `call`, C# an `()` operator — one slot, so a callable
                // object stays callable across the language boundary.
                let dunder_prop =
                    self.str_const(&vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Call));
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
            if let Some(callable_global) =
                self.source_function_callable_global_name_for_canon(&canon_name)
            {
                let fn_idx = self.str_const(&callable_global);
                self.emit_u16(Op::GLOBAL_GET, fn_idx);
            } else if !is_local {
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
            self.emit_source_function_callable_name_resolution(callee_slot);
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
                let rest_receiver = if self.profile.name == "lua" {
                    None
                } else {
                    Some(receiver_slot)
                };
                self.emit_call_ref_with_arg_slots(callee_slot, rest_receiver, &arg_slots);
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
        if self.class_prototype_dispatch() {
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
                // §13.3.7: the property key and the arguments are evaluated with
                // the CALLER's `this`; `this` is bound to the receiver only for
                // the call itself (just before CALL_REF below). Binding it here
                // would make `obj[k](this, …)` see the receiver, not the
                // enclosing `this`.
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
                let line = self.line;
                self.chunk().emit_if(line);
                // Direct ecma `[[Get]]` (Reflect.get) → callee_tmp; a non-object
                // receiver yields undefined (Reflect.get throws on non-object,
                // where the old __vybe_js_get_method returned undefined).
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                crate::primitives::instructions::recipes::is_object(self.chunk(), line);
                self.chunk().emit_if_value(line);
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
                let reflect_idx = self.import("ecma:reflect", "get");
                self.emit_host_call(reflect_idx, 2);
                self.chunk().emit_else(line);
                inst!(self, core_wasm::undefined);
                self.chunk().emit_end(line);
                self.emit_u16(Op::LOCAL_SET, callee_tmp);
                self.chunk().emit_else(line);
                self.emit_u16(Op::LOCAL_GET, callee_tmp);
                fn_call!(self, "wasm:js-undefined", "test", 1);
                let line = self.line;
                self.chunk().emit_if(line);
                // Direct ecma `[[Get]]` (Reflect.get) → callee_tmp; a non-object
                // receiver yields undefined (Reflect.get throws on non-object,
                // where the old __vybe_js_get_method returned undefined).
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                crate::primitives::instructions::recipes::is_object(self.chunk(), line);
                self.chunk().emit_if_value(line);
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
                let reflect_idx = self.import("ecma:reflect", "get");
                self.emit_host_call(reflect_idx, 2);
                self.chunk().emit_else(line);
                inst!(self, core_wasm::undefined);
                self.chunk().emit_end(line);
                self.emit_u16(Op::LOCAL_SET, callee_tmp);
                self.chunk().emit_end(line);
                self.chunk().emit_end(line);

                self.emit_u16(Op::LOCAL_GET, callee_tmp);
                for a in &arg_exprs {
                    self.compile_expr(a)?;
                }
                // Bind `this` = receiver for the call only. Stack is
                // [callee, ..args, receiver]; GLOBAL_SET pops the receiver,
                // leaving [callee, ..args] for CALL_REF.
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.set_js_this_from_stack();
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
        self.emit_source_function_callable_name_resolution(callee_slot);

        // A callee that is an OBJECT, not a function: dispatch through the Call
        // SLOT. `Counter()(3)`, `A()(3)(4)`, `makeAdder()(2)` — the callee here
        // is a produced value, so the Ident path's probe never sees it and a
        // plain `CALL_REF` trapped with "Not a function".
        //
        // By ROLE, never by spelling: Kotlin `operator fun invoke`, Python
        // `__call__`, PHP `__invoke`, Dart `call` and a C# `()` operator all
        // fill `ProtocolSlot::Call`, so one probe reaches every one of them and
        // no method-name table appears in shared code. Rewrites `callee_slot`
        // to the bound method and remembers the object as the receiver, so the
        // existing dispatch below carries it as `this`.
        let invoke_matched = self.define_local("__call_ref_invoke_matched");
        let invoke_receiver = self.define_local("__call_ref_invoke_receiver");
        self.emit_const(Value::I32(0));
        self.emit_u16(Op::LOCAL_SET, invoke_matched);
        if self.profile.callable_objects {
            let line = self.line;
            // STRUCT_GET traps on a primitive, so gate on it being an object.
            self.emit_u16(Op::LOCAL_GET, callee_slot);
            let typeof_idx = self.import("ecma:value", "typeof");
            self.emit_host_call(typeof_idx, 1);
            self.emit_const(Value::String(Arc::from("object")));
            crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
            self.chunk().emit_if(line);

            self.emit_u16(Op::LOCAL_GET, callee_slot);
            self.emit_u16(Op::LOCAL_SET, invoke_receiver);
            self.emit_u16(Op::LOCAL_GET, callee_slot);
            let slot_key =
                self.str_const(&vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Call));
            self.emit_u16(Op::STRUCT_GET, slot_key);
            let method_slot = self.define_local("__call_ref_invoke_method");
            self.emit_u16(Op::LOCAL_SET, method_slot);

            self.emit_u16(Op::LOCAL_GET, method_slot);
            self.emit(Op::REF_IS_NULL);
            self.emit(Op::I32_EQZ);
            let line = self.line;
            self.chunk().emit_if(line);
            self.emit_u16(Op::LOCAL_GET, method_slot);
            self.emit_u16(Op::LOCAL_SET, callee_slot);
            self.emit_const(Value::I32(1));
            self.emit_u16(Op::LOCAL_SET, invoke_matched);
            self.chunk().emit_end(line);

            self.chunk().emit_end(line);
        }

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

        // The Call-slot method is UNBOUND — its receiver is the object the call
        // was written on, which `__vybe_method_receiver` on the method itself
        // does not carry.
        {
            let line = self.line;
            self.emit_u16(Op::LOCAL_GET, invoke_matched);
            self.chunk().emit_if(line);
            self.emit_u16(Op::LOCAL_GET, invoke_receiver);
            self.emit_u16(Op::LOCAL_SET, receiver_slot);
            self.chunk().emit_end(line);
        }

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
}
