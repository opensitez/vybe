//! Call-expression compilation — `compile_call` (handles named calls,
//! method calls, super-calls, spread, dotted lookups) and
//! `compile_lambda`. This is the primary edit site for the inline
//! refactor (Phase G) where `wasm:js-*` imports get replaced by
//! inline WASM GC sequences.

use super::*;
use crate::scope::UpvalueDesc;

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
        && matches!(field.as_str(), "FromDays" | "FromHours" | "FromMinutes" | "FromSeconds" | "FromMilliseconds" | "Zero")
    {
        return Some("TimeSpan".into());
    }
    if class_name.eq_ignore_ascii_case("DateTime")
        && matches!(field.as_str(), "Now" | "UtcNow" | "Today" | "Parse")
    {
        return Some("DateTime".into());
    }
    if class_name.eq_ignore_ascii_case("Convert")
        && field.eq_ignore_ascii_case("ToDateTime")
    {
        return Some("DateTime".into());
    }
    if class_name.eq_ignore_ascii_case("Guid")
        && matches!(field.as_str(), "Empty" | "NewGuid" | "Parse")
    {
        return Some("Guid".into());
    }
    if class_name.eq_ignore_ascii_case("Version")
        && matches!(field.as_str(), "Parse")
    {
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
    if class_name.eq_ignore_ascii_case("TimeSpan")
        && field == "Zero"
    {
        return Some("TimeSpan".into());
    }
    if class_name.eq_ignore_ascii_case("Guid")
        && field == "Empty"
    {
        return Some("Guid".into());
    }
    if class_name.eq_ignore_ascii_case("Version")
        && field == "Parse"
    {
        return Some("Version".into());
    }
    None
}

fn resolve_receiver_type_hint(compiler: &Compiler, recv: &Expression) -> Option<String> {
    match &recv.kind {
        ExprKind::Ident(local_name) => compiler.lookup_var_type_hint(local_name).map(str::to_string)
            .or_else(|| compiler.scope().resolve_type_ci(local_name).map(|s| s.to_string()))
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
                let owner_name = owner
                    .split('<')
                    .next()
                    .map(str::trim)
                    .unwrap_or(owner);
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
        ExprKind::New { class, .. } => terminal_type_name(class)
            .map(|name| compiler.resolve_source_type_alias(&name)),
        ExprKind::Call { callee, args, .. } => {
            let arg_exprs: Vec<&Expression> = args.iter().map(|arg| &arg.value).collect();
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                if let Some(return_type) = compiler
                    .resolve_instance_method_overload(object, field, &arg_exprs, false)
                    .and_then(|overload| overload.return_type.clone())
                {
                    return Some(return_type);
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
                        if let Some(class_name) = compiler.resolve_pending_class_name_for_type_hint(&receiver_type) {
                            if compiler.pending_classes.get(&class_name).is_some_and(|pending| {
                                pending
                                    .instance_pointer_method_names
                                    .iter()
                                    .any(|name| compiler.canon(name) == compiler.canon(field))
                            }) {
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
    if compiler.scope().resolve(head_name).is_some()
        || compiler.scope().resolve_ci(head_name).is_some()
        || compiler.lookup_var_type_hint(head_name).is_some()
    {
        return false;
    }

    let full_canon = compiler.canon(&class_parts.join("."));
    let short_canon = compiler.canon(class_parts.last().map(String::as_str).unwrap_or(""));
    let method_canon = compiler.canon(field);

    [full_canon, short_canon]
        .into_iter()
        .any(|container_canon| {
            compiler.defined_classes.contains(&container_canon)
                && compiler
                    .pending_classes
                    .get(container_canon.as_str())
                    .map(|pending| pending.static_method_names.iter().any(|name| name == &method_canon))
                    .unwrap_or(false)
        })
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
    fn resolve_php_autoload_callback_class_global(&self, class_name: &str) -> Option<String> {
        let resolved_class = self.resolve_source_type_alias(class_name);
        let canon_class = self.canon(&resolved_class);
        if self.defined_classes.contains(&canon_class) || self.defined_globals.contains(&canon_class) {
            return Some(canon_class);
        }
        resolved_class.rsplit('.').next().and_then(|short_name| {
            let short_canon = self.canon(short_name);
            if self.defined_classes.contains(&short_canon) || self.defined_globals.contains(&short_canon) {
                Some(short_canon)
            } else {
                None
            }
        })
    }

    fn compile_php_autoload_callable_ref(&mut self, expr: &Expression) -> Result<(), String> {
        match &expr.kind {
            ExprKind::Lit(Literal::Str(function_name)) => {
                let resolved_name = self.resolve_source_type_alias(function_name);
                let function_idx = self.str_const(&self.canon(&resolved_name));
                self.emit_u16(Op::GLOBAL_GET, function_idx);
                Ok(())
            }
            _ => self.compile_expr(expr),
        }
    }

    fn overload_type_matches(&self, param_type: &str, arg_type: &str) -> bool {
        let normalized_param = Self::normalize_type_hint(strip_generic_suffix(param_type).trim_end_matches('?'));
        let normalized_arg = Self::normalize_type_hint(strip_generic_suffix(arg_type).trim_end_matches('?'));
        normalized_param == normalized_arg
            || (Self::is_string_type_hint(&normalized_param) && Self::is_string_type_hint(&normalized_arg))
            || (matches!(normalized_param.as_str(), "bool" | "boolean")
                && matches!(normalized_arg.as_str(), "bool" | "boolean"))
            || (is_numeric_overload_type(&normalized_param) && is_numeric_overload_type(&normalized_arg))
    }

    fn match_method_overload(
        &self,
        overloads: &[PendingMethodOverload],
        arg_exprs: &[&Expression],
        require_multiple: bool,
    ) -> Option<PendingMethodOverload> {
        if require_multiple && overloads.len() < 2 {
            return None;
        }

        let same_arity: Vec<&PendingMethodOverload> = overloads
            .iter()
            .filter(|overload| overload.param_types.len() == arg_exprs.len())
            .collect();
        if same_arity.len() == 1 {
            return Some(same_arity[0].clone());
        }

        let arg_types: Vec<Option<String>> = arg_exprs
            .iter()
            .map(|expr| self.infer_expr_type_hint(expr).map(|hint| Self::normalize_type_hint(&hint)))
            .collect();
        if arg_types.iter().any(|hint| hint.is_none()) {
            return None;
        }

        let matching: Vec<&PendingMethodOverload> = same_arity
            .into_iter()
            .filter(|overload| {
                overload
                    .param_types
                    .iter()
                    .zip(arg_types.iter())
                    .all(|(param_type, arg_type)| {
                        arg_type
                            .as_deref()
                            .is_some_and(|arg_type| self.overload_type_matches(param_type, arg_type))
                    })
            })
            .collect();
        if matching.len() == 1 {
            Some(matching[0].clone())
        } else {
            None
        }
    }

    fn match_method_overload_chunk(
        &self,
        overloads: &[PendingMethodOverload],
        arg_exprs: &[&Expression],
        require_multiple: bool,
    ) -> Option<usize> {
        self.match_method_overload(overloads, arg_exprs, require_multiple)
            .map(|overload| overload.chunk_idx)
    }

    fn resolve_instance_method_overload(
        &self,
        object: &Expression,
        field: &str,
        arg_exprs: &[&Expression],
        require_multiple: bool,
    ) -> Option<PendingMethodOverload> {
        let receiver_type = resolve_receiver_type_hint(self, object)?;
        let class_name = self.resolve_pending_class_name_for_type_hint(&receiver_type)?;
        let pending = self.pending_classes.get(&class_name)?;
        let method_key = self.canon(field);
        let overloads = pending.instance_method_overloads.get(&method_key)?;
        self.match_method_overload(overloads, arg_exprs, require_multiple)
    }

    fn resolve_instance_method_overload_chunk(
        &self,
        object: &Expression,
        field: &str,
        arg_exprs: &[&Expression],
    ) -> Option<usize> {
        self.resolve_instance_method_overload(object, field, arg_exprs, true)
            .map(|overload| overload.chunk_idx)
    }

    pub(super) fn resolve_pending_class_name_for_type_hint(&self, type_hint: &str) -> Option<String> {
        let resolved_type = self.resolve_source_type_alias(type_hint);
        let receiver_type = resolved_type
            .trim()
            .trim_end_matches('?')
            .trim_start_matches('*')
            .trim_start_matches('^')
            .trim();
        let receiver_canon = self.canon(strip_generic_suffix(receiver_type));
        if self.pending_classes.contains_key(&receiver_canon) {
            return Some(receiver_canon);
        }
        self.pending_classes
            .keys()
            .find(|name| name.eq_ignore_ascii_case(receiver_type) || name.eq_ignore_ascii_case(&receiver_canon))
            .cloned()
    }

    pub(super) fn pending_class_has_method_name_for_type(&self, type_hint: &str, method_name: &str) -> bool {
        let Some(class_name) = self.resolve_pending_class_name_for_type_hint(type_hint) else {
            return false;
        };
        let Some(pending) = self.pending_classes.get(&class_name) else {
            return false;
        };
        let method_key = self.canon(method_name);
        pending.static_method_overloads.contains_key(&method_key)
            || pending.instance_method_overloads.contains_key(&method_key)
            || pending.static_method_names.iter().any(|name| self.canon(name) == method_key)
            || pending.instance_member_names.iter().any(|name| self.canon(name) == method_key)
    }

    fn direct_receiver_has_own_pending_method(&self, receiver: &Expression, method_name: &str) -> bool {
        let class_name = match &receiver.kind {
            ExprKind::This | ExprKind::Super => self.current_class.clone(),
            ExprKind::Ident(name) => {
                let canon = self.canon(name);
                if canon == self.profile.self_keyword || canon == "me" || canon == "this" || canon == "mybase" {
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

        let method_key = self.canon(method_name);
        pending.instance_method_overloads.contains_key(&method_key)
            || pending.instance_member_names.iter().any(|name| self.canon(name) == method_key)
    }

    pub(super) fn resolve_static_method_overload_chunk_for_type(
        &self,
        type_hint: &str,
        method_name: &str,
        arg_exprs: &[&Expression],
    ) -> Option<usize> {
        let class_name = self.resolve_pending_class_name_for_type_hint(type_hint)?;
        let pending = self.pending_classes.get(&class_name)?;
        let method_key = self.canon(method_name);
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
        let class_name = self.resolve_pending_class_name_for_type_hint(type_hint)?;
        let pending = self.pending_classes.get(&class_name)?;
        let method_key = self.canon(method_name);
        let overloads = pending
            .static_method_overloads
            .get(&method_key)
            .or_else(|| pending.instance_method_overloads.get(&method_key))?;
        self.match_method_overload(overloads, arg_exprs, false)
    }

    fn emit_direct_instance_method_call(
        &mut self,
        chunk_idx: usize,
        obj_tmp: u16,
        arg_exprs: &[&Expression],
    ) -> Result<(), String> {
        let line = self.line;
        self.emit_u16(Op::REF_FUNC, chunk_idx as u16);
        self.chunk().emit(0, line);
        let fn_tmp = self.define_local("__direct_instance_method_fn");
        self.emit_u16(Op::LOCAL_SET, fn_tmp);
        self.emit(Op::DROP);
        let mut arg_slots = Vec::with_capacity(arg_exprs.len());
        for (index, arg) in arg_exprs.iter().enumerate() {
            self.compile_expr(arg)?;
            let arg_slot = self.define_local(&format!("__direct_instance_method_arg_{}", index));
            self.emit_u16(Op::LOCAL_SET, arg_slot);
            self.emit(Op::DROP);
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
        self.emit(Op::DROP);
        let mut arg_slots = Vec::with_capacity(arg_exprs.len());
        for (index, arg) in arg_exprs.iter().enumerate() {
            self.compile_expr(arg)?;
            let arg_slot = self.define_local(&format!("__direct_static_method_arg_{}", index));
            self.emit_u16(Op::LOCAL_SET, arg_slot);
            self.emit(Op::DROP);
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
        self.emit(Op::DUP);
        self.emit_const(Value::F64(fixed_count as f64));
        self.emit_u16(Op::STRUCT_SET, key);
        self.emit(Op::DROP);
    }

    fn emit_normal_call_from_arg_slots(&mut self, callee_slot: u16, receiver_slot: Option<u16>, arg_slots: &[u16]) {
        self.emit_u16(Op::LOCAL_GET, callee_slot);
        if let Some(receiver_slot) = receiver_slot {
            self.emit_u16(Op::LOCAL_GET, receiver_slot);
        }
        for slot in arg_slots {
            self.emit_u16(Op::LOCAL_GET, *slot);
        }
        self.emit_u8(Op::CALL_REF, (arg_slots.len() + usize::from(receiver_slot.is_some())) as u8);
    }

    fn emit_rest_call_from_arg_slots(
        &mut self,
        callee_slot: u16,
        receiver_slot: Option<u16>,
        arg_slots: &[u16],
        fixed_count: usize,
    ) {
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
                self.emit(Op::UNDEFINED);
            }
        }
        let rest_slot = self.define_local("__runtime_rest_call_array");
        common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
        self.emit_u16(Op::LOCAL_SET, rest_slot);
        self.emit(Op::DROP);
        for slot in arg_slots.iter().skip(fixed_count) {
            self.emit_u16(Op::LOCAL_GET, rest_slot);
            self.emit_u16(Op::LOCAL_GET, *slot);
            common::collections::emit_push(&mut self.chunks, self.current, line);
            self.emit(Op::DROP);
        }
        self.emit_u16(Op::LOCAL_GET, rest_slot);
        self.emit_u8(Op::CALL_REF, argc as u8);
    }

    fn emit_array_value_or_undefined(&mut self, args_slot: u16, len_slot: u16, index: usize) {
        self.emit_u16(Op::LOCAL_GET, len_slot);
        self.emit_const(Value::F64(index as f64));
        self.emit(Op::DYN_GT);
        let has_value = self.emit_jump(Op::BR_IF_FALSE);
        self.emit_u16(Op::LOCAL_GET, args_slot);
        self.emit_const(Value::F64(index as f64));
        common::collections::emit_get(&mut self.chunks, self.current, self.line);
        let done = self.emit_jump(Op::BR);
        self.patch_jump(has_value);
        self.emit(Op::UNDEFINED);
        self.patch_jump(done);
    }

    fn emit_normal_call_from_args_array(&mut self, callee_slot: u16, receiver_slot: Option<u16>, args_slot: u16, known_len: Option<usize>) {
        let line = self.line;
        if let Some(known_len) = known_len {
            self.emit_u16(Op::LOCAL_GET, callee_slot);
            if let Some(receiver_slot) = receiver_slot {
                self.emit_u16(Op::LOCAL_GET, receiver_slot);
            }
            self.emit_u16(Op::LOCAL_GET, args_slot);
            self.emit(Op::SPREAD);
            self.emit_u8(Op::CALL_REF, (known_len + usize::from(receiver_slot.is_some())) as u8);
            return;
        }

        self.emit_u16(Op::LOCAL_GET, args_slot);
        self.emit_const(Value::I32(16));
        common::collections::emit_new_with_length(&mut self.chunks, self.current, line);
        common::collections::emit_concat(&mut self.chunks, self.current, line);
        self.emit_const(Value::F64(0.0));
        self.emit_const(Value::F64(16.0));
        common::collections::emit_slice(&mut self.chunks, self.current, line);
        let padded_slot = self.define_local("__runtime_spread_args16");
        self.emit_u16(Op::LOCAL_SET, padded_slot);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, callee_slot);
        if let Some(receiver_slot) = receiver_slot {
            self.emit_u16(Op::LOCAL_GET, receiver_slot);
        }
        self.emit_u16(Op::LOCAL_GET, padded_slot);
        self.emit(Op::SPREAD);
        self.emit_u8(Op::CALL_REF, (16 + usize::from(receiver_slot.is_some())) as u8);
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
                        self.emit(Op::UNDEFINED);
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
                self.emit(Op::DROP);
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
        arg_slots: &[u16],
        result_slot: u16,
    ) {
        let rest_fixed_counts: Vec<u8> = self.rest_fixed_arities.iter().copied().collect();
        if rest_fixed_counts.is_empty() {
            self.emit_normal_call_from_arg_slots(callee_slot, receiver_slot, arg_slots);
            self.emit_u16(Op::LOCAL_SET, result_slot);
            self.emit(Op::DROP);
            return;
        }

        let rest_key = self.str_const("__vybe_rest_fixed_arity");
        let rest_arity_slot = self.define_local("__call_rest_fixed_arity");
        self.emit_u16(Op::LOCAL_GET, callee_slot);
        self.emit_u16(Op::STRUCT_GET, rest_key);
        self.emit_u16(Op::LOCAL_SET, rest_arity_slot);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, rest_arity_slot);
        self.emit(Op::REF_IS_NULL);
        let no_rest = self.emit_jump(Op::BR_IF_TRUE);
        let mut rest_done = Vec::new();
        for fixed_count in rest_fixed_counts {
            self.emit_u16(Op::LOCAL_GET, rest_arity_slot);
            self.emit_const(Value::F64(fixed_count as f64));
            self.emit(Op::DYN_EQ);
            let next = self.emit_jump(Op::BR_IF_FALSE);
            self.emit_rest_call_from_arg_slots(callee_slot, receiver_slot, arg_slots, fixed_count as usize);
            self.emit_u16(Op::LOCAL_SET, result_slot);
            self.emit(Op::DROP);
            rest_done.push(self.emit_jump(Op::BR));
            self.patch_jump(next);
        }
        self.patch_jump(no_rest);
        self.emit_normal_call_from_arg_slots(callee_slot, receiver_slot, arg_slots);
        self.emit_u16(Op::LOCAL_SET, result_slot);
        self.emit(Op::DROP);
        for done in rest_done {
            self.patch_jump(done);
        }
    }

    fn emit_call_ref_with_arg_slots(&mut self, callee_slot: u16, receiver_slot: Option<u16>, arg_slots: &[u16]) {
        let result_slot = self.define_local("__call_runtime_result");
        if let Some(receiver_slot) = receiver_slot {
            self.emit_u16(Op::LOCAL_GET, receiver_slot);
            self.emit(Op::REF_IS_NULL);
            let no_receiver = self.emit_jump(Op::BR_IF_TRUE);
            self.emit_dispatch_and_store_from_arg_slots(callee_slot, Some(receiver_slot), arg_slots, result_slot);
            let done = self.emit_jump(Op::BR);
            self.patch_jump(no_receiver);
            self.emit_dispatch_and_store_from_arg_slots(callee_slot, None, arg_slots, result_slot);
            self.patch_jump(done);
        } else {
            self.emit_dispatch_and_store_from_arg_slots(callee_slot, None, arg_slots, result_slot);
        }
        self.emit_u16(Op::LOCAL_GET, result_slot);
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
            self.emit(Op::DROP);

            self.emit_u16(Op::LOCAL_GET, rest_arity_slot);
            self.emit(Op::REF_IS_NULL);
            let no_rest = self.emit_jump(Op::BR_IF_TRUE);
            let mut rest_done = Vec::new();
            for fixed_count in rest_fixed_counts {
                self.emit_u16(Op::LOCAL_GET, rest_arity_slot);
                self.emit_const(Value::F64(fixed_count as f64));
                self.emit(Op::DYN_EQ);
                let next = self.emit_jump(Op::BR_IF_FALSE);
                self.emit_rest_call_from_args_array(callee_slot, receiver_slot, args_slot, known_len, fixed_count as usize);
                self.emit_u16(Op::LOCAL_SET, result_slot);
                self.emit(Op::DROP);
                rest_done.push(self.emit_jump(Op::BR));
                self.patch_jump(next);
            }
            self.patch_jump(no_rest);
            self.emit_normal_call_from_args_array(callee_slot, receiver_slot, args_slot, known_len);
            self.emit_u16(Op::LOCAL_SET, result_slot);
            self.emit(Op::DROP);
            for done in rest_done {
                self.patch_jump(done);
            }
            return;
        }

        self.emit_normal_call_from_args_array(callee_slot, receiver_slot, args_slot, known_len);
        self.emit_u16(Op::LOCAL_SET, result_slot);
        self.emit(Op::DROP);
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
            let no_receiver = self.emit_jump(Op::BR_IF_TRUE);
            self.emit_dispatch_and_store_from_args_array(callee_slot, Some(receiver_slot), args_slot, known_len, result_slot);
            let done = self.emit_jump(Op::BR);
            self.patch_jump(no_receiver);
            self.emit_dispatch_and_store_from_args_array(callee_slot, None, args_slot, known_len, result_slot);
            self.patch_jump(done);
        } else {
            self.emit_dispatch_and_store_from_args_array(callee_slot, None, args_slot, known_len, result_slot);
        }
        self.emit_u16(Op::LOCAL_GET, result_slot);
    }

    fn emit_php_dynamic_function_name_resolution(&mut self, callee_slot: u16) {
        if !self.is_php_profile() {
            return;
        }

        let mut known_functions: Vec<String> = self.defined_functions.iter().cloned().collect();
        if known_functions.is_empty() {
            return;
        }
        known_functions.sort();

        self.emit_u16(Op::LOCAL_GET, callee_slot);
        self.emit(Op::REF_TYPEOF);
        self.emit_const(Value::String(Arc::from("string")));
        self.emit(Op::DYN_EQ);
        let not_string = self.emit_jump(Op::BR_IF_FALSE);

        self.emit_u16(Op::LOCAL_GET, callee_slot);
    let line = self.line;
    common::strings::emit_to_lower(self.chunk(), line);
        let callee_name_slot = self.define_local("__php_string_callee_name");
        self.emit_u16(Op::LOCAL_SET, callee_name_slot);
        self.emit(Op::DROP);

        let mut done_jumps = Vec::new();
        for function_name in known_functions {
            let lowered_name = function_name.to_ascii_lowercase();
            self.emit_u16(Op::LOCAL_GET, callee_name_slot);
            self.emit_const(Value::String(Arc::from(lowered_name.as_str())));
            self.emit(Op::DYN_EQ);
            let next = self.emit_jump(Op::BR_IF_FALSE);

            let idx = self.str_const(&function_name);
            self.emit_u16(Op::GLOBAL_GET, idx);
            self.emit_u16(Op::LOCAL_SET, callee_slot);
            self.emit(Op::DROP);
            done_jumps.push(self.emit_jump(Op::BR));

            self.patch_jump(next);
        }

        self.patch_jump(not_string);
        for done in done_jumps {
            self.patch_jump(done);
        }
    }

    fn emit_flat_call_args_array(&mut self, args: &[Argument], slot_name: &str) -> Result<u16, String> {
        let line = self.line;
        let args_slot = self.define_local(slot_name);
        common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
        self.emit_u16(Op::LOCAL_SET, args_slot);
        self.emit(Op::DROP);
        for arg in args {
            if arg.spread {
                self.emit_u16(Op::LOCAL_GET, args_slot);
                self.compile_expr(&arg.value)?;
                common::collections::emit_concat(&mut self.chunks, self.current, line);
                self.emit_u16(Op::LOCAL_SET, args_slot);
                self.emit(Op::DROP);
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
                    self.emit(Op::UNDEFINED);
                }
            }

            let line = self.line;
            let rest_slot = self.define_local("__packed_rest_call_array");
            common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
            self.emit_u16(Op::LOCAL_SET, rest_slot);
            self.emit(Op::DROP);
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
        self.emit(Op::DROP);

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

    pub(super) fn emit_js_exception_ctor_from_message_value(&mut self, type_name: &str) -> Result<(), String> {
        let msg_val = self.define_local("__exc_msg_val");
        self.emit_u16(Op::LOCAL_SET, msg_val);
        self.emit(Op::DROP);

        self.emit_u16(Op::STRUCT_NEW, 0);
        self.emit(Op::DUP);
        self.emit_u16(Op::LOCAL_GET, msg_val);
        let line = self.line;
        common::errors::emit_exception_new_finalize(self.chunk(), type_name, line);

        let exc_tmp = self.define_local("__exc_tmp");
        self.emit_u16(Op::LOCAL_SET, exc_tmp);
        self.emit(Op::DROP);

        self.emit_const(Value::String(Arc::from(format!("{}: ", type_name))));
        self.emit_u16(Op::LOCAL_GET, exc_tmp);
        let msg_k = self.str_const("message");
        self.emit_u16(Op::STRUCT_GET, msg_k);
        self.emit(Op::STR_CONCAT);
        let stack_val = self.define_local("__stack_val");
        self.emit_u16(Op::LOCAL_SET, stack_val);
        self.emit(Op::DROP);
        self.emit_u16(Op::LOCAL_GET, exc_tmp);
        self.emit_u16(Op::LOCAL_GET, stack_val);
        let stack_key = self.str_const("stack");
        self.emit_u16(Op::STRUCT_SET, stack_key);
        self.emit(Op::DROP);

        if self.is_js_profile() {
            for name in Self::js_error_instanceof_chain(type_name) {
                common::classes::emit_instanceof_chain(&mut self.chunks, self.current, exc_tmp, name, line);
            }
        }

        self.emit_u16(Op::LOCAL_GET, exc_tmp);
        Ok(())
    }

    pub(super) fn emit_js_exception_ctor_value(&mut self, type_name: &str, args: &[&Expression]) -> Result<(), String> {
        if let Some(msg_arg) = args.first() {
            self.compile_expr(msg_arg)?;
        } else {
            self.emit_const(Value::String(Arc::from("")));
        }
        self.emit_js_exception_ctor_from_message_value(type_name)?;

        if let Some(opts_arg) = args.get(1) {
            let exc_tmp = self.define_local("__exc_with_cause");
            self.emit_u16(Op::LOCAL_SET, exc_tmp);
            self.emit(Op::DROP);
            self.compile_expr(opts_arg)?;
            let cause_key = self.str_const("cause");
            self.emit_u16(Op::STRUCT_GET, cause_key);
            let cause_val = self.define_local("__cause_val");
            self.emit_u16(Op::LOCAL_SET, cause_val);
            self.emit(Op::DROP);
            self.emit_u16(Op::LOCAL_GET, exc_tmp);
            self.emit_u16(Op::LOCAL_GET, cause_val);
            self.emit_u16(Op::STRUCT_SET, cause_key);
            self.emit(Op::DROP);
            self.emit_u16(Op::LOCAL_GET, exc_tmp);
        }
        Ok(())
    }

    pub(super) fn emit_generator_control_packet_from_stack(&mut self, op: &str) {
        let value_slot = self.define_local("__gen_control_value");
        self.emit_u16(Op::LOCAL_SET, value_slot); self.emit(Op::DROP);

        let line = self.line;
        common::dict::emit_new(&mut self.chunks, self.current, line);

        self.emit(Op::DUP);
        self.emit_const(Value::Bool(true));
        let marker_key = self.str_const("__vybe_generator_control");
        self.emit_u16(Op::STRUCT_SET, marker_key);
        self.emit(Op::DROP);

        self.emit(Op::DUP);
        self.emit_const(Value::String(Arc::from(op)));
        let op_key = self.str_const("op");
        self.emit_u16(Op::STRUCT_SET, op_key);
        self.emit(Op::DROP);

        self.emit(Op::DUP);
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

    fn try_compile_go_map_has_call(&mut self, callee: &Expression, args: &[Argument]) -> Result<bool, String> {
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
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, map_slot);
        self.emit(Op::REF_IS_NULL);
        let non_null = self.emit_jump(Op::BR_IF_FALSE);
        self.emit(Op::FALSE);
        let end = self.emit_jump(Op::BR);

        self.patch_jump(non_null);
        self.emit_u16(Op::LOCAL_GET, map_slot);
        self.compile_expr(&args[1].value)?;
        let line = self.line;
        common::dict::emit_method_has(&mut self.chunks, self.current, line);
        self.patch_jump(end);
        Ok(true)
    }

    pub(super) fn compile_call(&mut self, callee: &Expression, args: &[Argument]) -> Result<(), String> {
        let reordered_args;
        let args = if args.iter().any(|arg| arg.name.is_some()) {
            reordered_args = self.reorder_named_call_args(callee, args);
            reordered_args.as_slice()
        } else {
            args
        };
        let arg_exprs: Vec<&Expression> = args.iter().map(|a| &a.value).collect();

        if self.try_compile_go_map_has_call(callee, args)? {
            return Ok(());
        }

        if let ExprKind::Member { object, field, null_safe } = &callee.kind {
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
                if let ExprKind::Call { callee: inner_callee, args: inner_args, .. } = &object.kind {
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
                            self.emit(Op::DUP);
                            self.compile_expr(&arg.value)?;
                            let key_idx = self.str_const(key);
                            self.emit_u16(Op::STRUCT_SET, key_idx);
                            self.emit(Op::DROP);

                            self.emit(Op::DUP);
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
                                let ExprKind::Tuple(items) = &element.value.kind else { continue; };
                                if items.len() != 2 { continue; }

                                self.emit(Op::DUP);
                                self.compile_expr(&items[0])?;
                                let key_tmp = self.define_local("__py_dict_ctor_key");
                                self.emit(Op::DUP);
                                self.emit_u16(Op::LOCAL_SET, key_tmp);
                                self.emit(Op::DROP);
                                self.compile_expr(&items[1])?;
                                common::collections::emit_set(&mut self.chunks, self.current, line);
                                self.emit(Op::DROP);

                                self.emit(Op::DUP);
                                let keys_key = self.str_const("__keys");
                                self.emit_u16(Op::STRUCT_GET, keys_key);
                                self.emit_u16(Op::LOCAL_GET, key_tmp);
                                common::collections::emit_push(&mut self.chunks, self.current, line);
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
                                if elements.len() == 2 && elements.iter().all(|element| element.key.is_none()) =>
                            {
                                let ExprKind::Lit(Literal::Str(class_name)) = &elements[0].value.kind else {
                                    self.emit(Op::UNDEFINED);
                                    self.emit_u16(Op::GLOBAL_SET, receiver_idx);
                                    self.emit(Op::DROP);
                                    self.compile_php_autoload_callable_ref(&callback.value)?;
                                    let global_idx = self.str_const("__php_autoload_callback");
                                    self.emit_u16(Op::GLOBAL_SET, global_idx);
                                    self.emit(Op::DROP);
                                    for arg in args.iter().skip(1) {
                                        self.compile_expr(&arg.value)?;
                                        self.emit(Op::DROP);
                                    }
                                    self.emit_const(Value::Bool(true));
                                    return Ok(());
                                };
                                let ExprKind::Lit(Literal::Str(method_name)) = &elements[1].value.kind else {
                                    self.emit(Op::UNDEFINED);
                                    self.emit_u16(Op::GLOBAL_SET, receiver_idx);
                                    self.emit(Op::DROP);
                                    self.compile_php_autoload_callable_ref(&callback.value)?;
                                    let global_idx = self.str_const("__php_autoload_callback");
                                    self.emit_u16(Op::GLOBAL_SET, global_idx);
                                    self.emit(Op::DROP);
                                    for arg in args.iter().skip(1) {
                                        self.compile_expr(&arg.value)?;
                                        self.emit(Op::DROP);
                                    }
                                    self.emit_const(Value::Bool(true));
                                    return Ok(());
                                };

                                if let Some(class_global) = self.resolve_php_autoload_callback_class_global(class_name) {
                                    let class_idx = self.str_const(&class_global);
                                    self.emit_u16(Op::GLOBAL_GET, class_idx);
                                    self.emit(Op::DUP);
                                    self.emit_u16(Op::GLOBAL_SET, receiver_idx);
                                    self.emit(Op::DROP);
                                    let method_idx = self.str_const(&self.canon(method_name));
                                    self.emit_u16(Op::STRUCT_GET, method_idx);
                                } else {
                                    self.emit(Op::UNDEFINED);
                                    self.emit_u16(Op::GLOBAL_SET, receiver_idx);
                                    self.emit(Op::DROP);
                                    self.compile_php_autoload_callable_ref(&callback.value)?;
                                }
                            }
                            _ => {
                                self.emit(Op::UNDEFINED);
                                self.emit_u16(Op::GLOBAL_SET, receiver_idx);
                                self.emit(Op::DROP);
                                self.compile_php_autoload_callable_ref(&callback.value)?;
                            }
                        }
                    } else {
                        self.emit(Op::UNDEFINED);
                        self.emit_u16(Op::GLOBAL_SET, receiver_idx);
                        self.emit(Op::DROP);
                        self.emit(Op::UNDEFINED);
                    }
                    let global_idx = self.str_const("__php_autoload_callback");
                    self.emit_u16(Op::GLOBAL_SET, global_idx);
                    self.emit(Op::DROP);

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

                    self.emit(Op::UNDEFINED);
                    let receiver_idx = self.str_const("__php_autoload_callback_receiver");
                    self.emit_u16(Op::GLOBAL_SET, receiver_idx);
                    self.emit(Op::DROP);
                    self.emit(Op::UNDEFINED);
                    let global_idx = self.str_const("__php_autoload_callback");
                    self.emit_u16(Op::GLOBAL_SET, global_idx);
                    self.emit(Op::DROP);
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
                        self.emit(Op::DUP);
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
                            let Some(key_expr) = &elem.key else { continue; };
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
                        self.emit(Op::DROP);

                        let count_slot = self.define_local("__php_extract_count");
                        self.emit_const(Value::I64(0));
                        self.emit_u16(Op::LOCAL_SET, count_slot);
                        self.emit(Op::DROP);

                        for bind_name in binding_names {
                            let key_name = bind_name.strip_prefix('$').unwrap_or(bind_name.as_str());
                            self.emit_u16(Op::LOCAL_GET, map_slot);
                            self.emit_const(Value::String(Arc::from(key_name)));
                            let line = self.line;
                            common::collections::emit_get(&mut self.chunks, self.current, line);
                            self.emit(Op::DUP);
                            self.emit(Op::REF_IS_NULL);
                            let skip_assign = self.emit_jump(Op::BR_IF_TRUE);

                            self.emit_var_set(&bind_name);
                            self.emit_u16(Op::LOCAL_GET, count_slot);
                            self.emit_const(Value::I64(1));
                            self.emit(Op::DYN_ADD);
                            self.emit_u16(Op::LOCAL_SET, count_slot);
                            self.emit(Op::DROP);
                            let after_assign = self.emit_jump(Op::BR);

                            self.patch_jump(skip_assign);
                            self.emit(Op::DROP);
                            self.patch_jump(after_assign);
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
                if let Some(parent_name) = self.pending_classes.get(class_name.as_str()).and_then(|pc| pc.parent.clone()) {
                    if common::errors::is_exception_type(&parent_name) {
                        self.emit_js_exception_ctor_value(&parent_name, &arg_exprs)?;
                        let self_kw = self.profile.self_keyword.clone();
                        if let Some(slot) = self.scope().resolve(&self_kw).or_else(|| self.scope().resolve_ci(&self_kw)) {
                            self.emit(Op::DUP);
                            self.emit_u16(Op::LOCAL_SET, slot);
                            self.emit(Op::DROP);
                        }
                        return Ok(());
                    }
                    let pname = self.canon(&parent_name);
                    let pidx = self.str_const(&pname);
                    self.emit_u16(Op::GLOBAL_GET, pidx);
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                    // Store result as this
                    let self_kw = self.profile.self_keyword.clone();
                    if let Some(slot) = self.scope().resolve(&self_kw).or_else(|| self.scope().resolve_ci(&self_kw)) {
                        self.emit(Op::DUP);
                        self.emit_u16(Op::LOCAL_SET, slot);
                        self.emit(Op::DROP);
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
                let parent_name = class_name.as_ref()
                    .and_then(|cn| self.pending_classes.get(cn.as_str()))
                    .and_then(|pc| pc.parent.clone());
                let self_kw = self.profile.self_keyword.clone();
                let self_slot = self.scope().resolve(&self_kw).or_else(|| self.scope().resolve_ci(&self_kw));

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
                        for a in &arg_exprs { self.compile_expr(a)?; }
                        self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                        let result_slot = self.define_local("__js_super_method_result");
                        self.emit_u16(Op::LOCAL_SET, result_slot);
                        self.emit(Op::DROP);
                        self.restore_js_this(saved_js_this);
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                    } else {
                        // Typed-language method ABI passes receiver as arg0.
                        if let Some(slot) = self_slot {
                            self.emit_u16(Op::LOCAL_GET, slot);
                        } else {
                            self.emit(Op::NULL);
                        }
                        for a in &arg_exprs { self.compile_expr(a)?; }
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
        // Available in all languages. Prints object properties to stderr.
        if let ExprKind::Ident(name) = &callee.kind {
            if name == "__debug_dump" {
                for a in &arg_exprs { self.compile_expr(a)?; }
                let idx = self.import("vybe:debug", "dump");
                self.emit_host_call(idx, arg_exprs.len() as u8);
                return Ok(());
            }

            let canon = self.canon(name);
            let shadows_builtin_exception = self.defined_functions.contains(&canon)
                || self.defined_classes.contains(&canon)
                || self.defined_globals.contains(&canon)
                || (!self.case_sensitive && (
                    self.defined_functions.iter().any(|g| g.eq_ignore_ascii_case(name))
                    || self.defined_classes.iter().any(|g| g.eq_ignore_ascii_case(name))
                    || self.defined_globals.iter().any(|g| g.eq_ignore_ascii_case(name))
                ));
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
                let class_name = Self::normalize_type_hint(&class_name);
                let surface = common::dotnet::surface();
                if let Some(target) = surface.lookup_instance_method(&class_name, field, arg_exprs.len() as u8) {
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
                                then: Box::new(Expression::new(ExprKind::Lit(Literal::Int(-1)))),
                                else_: Box::new(Expression::new(ExprKind::Ternary {
                                    cond: Box::new(Expression::new(ExprKind::Binary {
                                        op: BinOp::Gt,
                                        left: Box::new(Expression::ident("left")),
                                        right: Box::new(Expression::ident("right")),
                                    })),
                                    then: Box::new(Expression::new(ExprKind::Lit(Literal::Int(1)))),
                                    else_: Box::new(Expression::new(ExprKind::Lit(Literal::Int(0)))),
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
                        && class_name.rsplit('.').next().is_some_and(|name| name.eq_ignore_ascii_case("List") || name.eq_ignore_ascii_case("ArrayList"))
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
                    for a in &arg_exprs { self.compile_expr(a)?; }
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

        // ── ESM host-module import binding ──────────────────────────
        //
        // `import { createServer } from "wasi:http"` binds
        // `createServer` locally. Calling it here emits a direct
        // `CALL_IMPORT` against the recorded (module, fn) pair — the
        // import statement itself is the compile-time declaration.
        if let ExprKind::Ident(name) = &callee.kind {
            let key = self.canon(name);
            if let Some((module, func)) = self.host_import_bindings.get(&key).cloned() {
                for a in &arg_exprs { self.compile_expr(a)?; }
                let idx = self.import(&module, &func);
                self.emit_host_call(idx, arg_exprs.len() as u8);
                return Ok(());
            }
        }

        if let ExprKind::Member { object, field, .. } = &callee.kind {
            if resolves_to_static_container_method(self, object, field) {
                self.compile_expr(object)?;
                let obj_tmp = self.define_local("__static_container_obj");
                self.emit_u16(Op::LOCAL_SET, obj_tmp);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                let method_idx = self.str_const(&self.canon(field));
                self.emit_u16(Op::STRUCT_GET, method_idx);
                let fn_tmp = self.define_local("__static_container_fn");
                self.emit_u16(Op::LOCAL_SET, fn_tmp);
                self.emit(Op::DROP);
                let class_canon = self.canon(&self.flatten_member_chain(object).join("."));
                if self.profile.name == "csharp" && args.len() == 1 && !args[0].spread {
                    if self.resolve_static_method_overload_for_type(&class_canon, field, &arg_exprs)
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
                        if self.profile.name == "php" { Some(obj_tmp) } else { None },
                        args,
                        signature,
                    )?;
                } else if self.profile.name == "php" {
                    let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                    for (index, arg) in arg_exprs.iter().enumerate() {
                        self.compile_expr(arg)?;
                        let arg_slot = self.define_local(&format!("__static_container_php_arg_{}", index));
                        self.emit_u16(Op::LOCAL_SET, arg_slot);
                        self.emit(Op::DROP);
                        arg_slots.push(arg_slot);
                    }
                    self.emit_call_ref_with_arg_slots(fn_tmp, Some(obj_tmp), &arg_slots);
                } else {
                    let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                    for (index, arg) in arg_exprs.iter().enumerate() {
                        self.compile_expr(arg)?;
                        let arg_slot = self.define_local(&format!("__static_container_arg_{}", index));
                        self.emit_u16(Op::LOCAL_SET, arg_slot);
                        self.emit(Op::DROP);
                        arg_slots.push(arg_slot);
                    }
                    self.emit_call_ref_with_arg_slots(fn_tmp, None, &arg_slots);
                }
                if args.iter().any(|arg| arg.by_ref) {
                    let pack_slot = self.define_local("__static_container_by_ref_pack");
                    self.emit_u16(Op::LOCAL_SET, pack_slot);
                    self.emit(Op::DROP);
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
                    && self.defined_functions.iter().any(|g| g.eq_ignore_ascii_case(name)));
            if !shadows_builtin && self.try_compile_builtin(name, &arg_exprs)? {
                return Ok(());
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
                if self.try_compile_builtin(&compound, &arg_exprs)? { return Ok(()); }

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
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    let idx = self.import(&module, field);
                    self.emit_host_call(idx, arg_exprs.len() as u8);
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
            if let ExprKind::Member { object: inner_obj, field: inner_field, .. } = &object.kind {
                if let ExprKind::Ident(prefix) = &inner_obj.kind {
                    let prefix_lc = self.canon(prefix);
                    if matches!(prefix_lc.as_str(), "vybe" | "wasi" | "wasm") {
                        let module = format!("{}:{}", prefix_lc, self.canon(inner_field));
                        for a in &arg_exprs { self.compile_expr(a)?; }
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

                if let Some(class_canon) = early_static_class_canon {
                    let cls_idx = self.str_const(&class_canon);
                    self.emit_u16(Op::GLOBAL_GET, cls_idx);
                    let method_canon = self.canon(&method_name);
                    let qualified_method = self.canon(&format!("{}.{}", class_canon, method_name));
                    let method_idx = self.str_const(&method_canon);
                    self.emit_u16(Op::STRUCT_GET, method_idx);
                    let fn_tmp = self.scope().resolve("__early_static_fn")
                        .unwrap_or_else(|| self.define_local("__early_static_fn"));
                    self.emit_u16(Op::LOCAL_SET, fn_tmp); self.emit(Op::DROP);

                    if let Some(param_modes) = self
                        .function_param_modes
                        .get(&qualified_method)
                        .cloned()
                        .or_else(|| self.function_param_modes.get(&method_canon).cloned())
                    {
                        if param_modes.iter().any(|mode| matches!(mode, PassBy::Ref | PassBy::Out)) {
                            let mut arg_slots = Vec::with_capacity(args.len());
                            for (index, arg) in args.iter().enumerate() {
                                match param_modes.get(index).copied().unwrap_or(PassBy::Value) {
                                    PassBy::Out => self.emit(Op::NULL),
                                    PassBy::Ref | PassBy::Const | PassBy::Value => {
                                        self.compile_expr_with_value_copy(&arg.value)?;
                                    }
                                }
                                let arg_slot = self.define_local(&format!("__early_static_call_arg_{}", index));
                                self.emit_u16(Op::LOCAL_SET, arg_slot);
                                self.emit(Op::DROP);
                                arg_slots.push(arg_slot);
                            }

                            self.emit_u16(Op::LOCAL_GET, fn_tmp);
                            for slot in &arg_slots {
                                self.emit_u16(Op::LOCAL_GET, *slot);
                            }
                            self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);

                            let pack_slot = self.define_local("__early_static_ref_call_pack");
                            self.emit_u16(Op::LOCAL_SET, pack_slot);
                            self.emit(Op::DROP);
                            let mut ref_out_index = 1usize;
                            for (index, arg) in args.iter().enumerate() {
                                if !matches!(param_modes.get(index), Some(PassBy::Ref | PassBy::Out)) {
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

                    if self.profile.name == "csharp" && args.len() == 1 && !args[0].spread {
                        if self.resolve_static_method_overload_for_type(&class_canon, &method_name, &arg_exprs)
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
                            let arg_slot = self.define_local(&format!("__early_static_arg_{}", index));
                            self.emit_u16(Op::LOCAL_SET, arg_slot);
                            self.emit(Op::DROP);
                            arg_slots.push(arg_slot);
                        }
                        if self.profile.name == "php" {
                            let cls_idx = self.str_const(&class_canon);
                            self.emit_u16(Op::GLOBAL_GET, cls_idx);
                            let receiver_slot = self.define_local("__early_static_receiver");
                            self.emit_u16(Op::LOCAL_SET, receiver_slot);
                            self.emit(Op::DROP);
                            self.emit_call_ref_with_arg_slots(fn_tmp, Some(receiver_slot), &arg_slots);
                        } else {
                            self.emit_call_ref_with_arg_slots(fn_tmp, None, &arg_slots);
                        }
                    }
                    if args.iter().any(|arg| arg.by_ref) {
                        let pack_slot = self.define_local("__early_static_by_ref_pack");
                        self.emit_u16(Op::LOCAL_SET, pack_slot);
                        self.emit(Op::DROP);
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

                // Use dotnet resolver when enabled
                if self.profile.namespaces.use_dotnet_resolver {
                    let skip_simple_instance_chain = if lower_parts.len() == 2 {
                        let head = &parts[0];
                        self.has_accessible_local_binding(head)
                            || self.defined_globals.contains(head)
                            || self.defined_globals.iter().any(|g| g.eq_ignore_ascii_case(head))
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
                    let field_set: std::collections::HashSet<String> = if let Some(ref cn) = self.current_class {
                        self.pending_classes.get(cn.as_str())
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
                            if is_user_class_for_local(name) { return false; }
                            accessible_locals.iter().any(|local| local == name || local.eq_ignore_ascii_case(name))
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
                                self.emit(Op::DROP);
                                self.emit_u16(Op::LOCAL_GET, resized_slot);
                                self.compile_assign_target(&args[0].value)?;
                                self.emit(Op::NULL);
                                return Ok(());
                            }

                            if emit.eq_ignore_ascii_case("dotnet.console_writeline") && arg_exprs.len() == 1 {
                                self.emit_dotnet_console_arg(arg_exprs[0])?;
                            } else {
                                for a in &arg_exprs { self.compile_expr(a)?; }
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
                                    ExprKind::Ident(name) => self.lookup_var_type_hint(name)
                                        .is_some_and(|hint| Self::normalize_type_hint(hint) == "char"),
                                    _ => false,
                                };
                                if is_char_like {
                                    self.compile_expr(arg_exprs[0])?;
                                    self.emit(Op::I32_CONST_0);
                                    self.emit(Op::STR_CHAR_CODE_AT);
                                    return Ok(());
                                }
                            }
                            for a in &arg_exprs { self.compile_expr(a)?; }
                            let idx = self.import(&module, &func);
                            self.emit_host_call(idx, arg_exprs.len() as u8);
                            return Ok(());
                        }
                        common::dotnet::DottedResolution::NamespaceAccess { parts: ns_parts } => {
                            // If any contiguous sub-window of the chain is a profile namespace
                            // constant (e.g. ["system","math","pi","tostring"] where "math.pi"
                            // is a constant), emit the constant and dispatch remaining as a
                            // value method. Namespace prefix before the constant is discarded.
                            if ns_parts.len() >= 2 {
                                let mut found_window: Option<(usize, usize)> = None;
                                'outer: for start in 0..ns_parts.len().saturating_sub(1) {
                                    for end in ((start + 2)..=ns_parts.len().saturating_sub(0)).rev() {
                                        if end > ns_parts.len() { continue; }
                                        let key = ns_parts[start..end].join(".");
                                        if self.profile.lookup_constant(&key).is_some() {
                                            found_window = Some((start, end));
                                            break 'outer;
                                        }
                                    }
                                }
                                if let Some((_const_start, const_end)) = found_window {
                                    let key = ns_parts[_const_start..const_end].join(".");
                                    let cv = self.profile.lookup_constant(&key).cloned().unwrap();
                                    match &cv {
                                        ConstantValue::Float(f) => self.emit_const(Value::F64(*f)),
                                        ConstantValue::Str(s) => self.emit_const(Value::String(Arc::from(s.as_str()))),
                                    }
                                    let remaining = ns_parts[const_end..].to_vec();
                                    if let Some(method_name) = remaining.first() {
                                        let argc = arg_exprs.len() as u8;
                                        let def = self.profile.lookup_value_method(method_name, argc).cloned();
                                        if let Some(def) = def {
                                            for a in &arg_exprs { self.compile_expr(a)?; }
                                            let line = self.line;
                                            match &def.emit {
                                                BuiltinEmit::Stdlib(name) => {
                                                    // For stdlib: func ref must be pushed BEFORE object.
                                                    // But object is already on stack. Save it to a temp.
                                                    let tmp = self.define_local("__const_val");
                                                    self.emit_u16(Op::LOCAL_SET, tmp); self.emit(Op::DROP);
                                                    let global_name = Self::stdlib_global_name(name);
                                                    let name_idx = self.str_const(&global_name);
                                                    self.emit_u16(Op::GLOBAL_GET, name_idx);
                                                    self.emit_u16(Op::LOCAL_GET, tmp);
                                                    for a in &arg_exprs { self.compile_expr(a)?; }
                                                    self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                                                }
                                                BuiltinEmit::HostCall(module, func) => {
                                                    let idx = self.import(module, func);
                                                    self.emit_host_call(idx, (arg_exprs.len() + 1) as u8);
                                                }
                                                BuiltinEmit::Common(name) => {
                                                    let name = name.clone();
                                                    self.emit_common(&name, (arg_exprs.len() + 1) as u8, line);
                                                }
                                                BuiltinEmit::Opcode(op_name) => {
                                                    self.emit_named_opcode(op_name);
                                                }
                                                _ => {
                                                    // Fallback: STRUCT_GET the method and call_ref
                                                    let idx = self.str_const(method_name);
                                                    self.emit_u16(Op::STRUCT_GET, idx);
                                                    self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                                                }
                                            }
                                        } else {
                                            // No value method — STRUCT_GET and call_ref
                                            let idx = self.str_const(method_name);
                                            self.emit_u16(Op::STRUCT_GET, idx);
                                            for a in &arg_exprs { self.compile_expr(a)?; }
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
                                self.emit(Op::DUP);
                                self.emit_u16(Op::STRUCT_GET, method_idx);
                                let fn_tmp = self.define_local("__ns_fn");
                                self.emit_u16(Op::LOCAL_SET, fn_tmp); self.emit(Op::DROP);
                                let obj_tmp = self.define_local("__ns_obj");
                                self.reserve_local_slot(obj_tmp);
                                self.emit_u16(Op::LOCAL_SET, obj_tmp); self.emit(Op::DROP);
                                self.emit_u16(Op::LOCAL_GET, fn_tmp);
                                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                                for a in &arg_exprs { self.compile_expr(a)?; }
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
                                for a in &arg_exprs { self.compile_expr(a)?; }
                                self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                            }
                            return Ok(());
                        }
                        common::dotnet::DottedResolution::InstanceMember { local, members } => {
                            // Intercept `parent.Controls.Add(child)` for GUI.
                            // The .NET WinForms surface is `Form.Controls.Add(ctrl)`,
                            // MAUI is `parent.Children.Add(ctrl)`, etc. — all
                            // resolve to the canonical gui emitter.
                            if members.len() >= 2 && members[members.len()-2] == "controls" && members[members.len()-1] == "add" {
                                let line = self.line;
                                let add_idx = self.import("vybe:gui", common::gui::HOST_FN_ADD_CHILD);
                                self.emit_var_get(&local);
                                for a in &arg_exprs { self.compile_expr(a)?; }
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
                                        common::threading::emit_thread_start(self.chunk(), line);
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
                    let func = if lower_parts.len() == 2 { lower_parts[1].clone() } else { lower_parts[1..].join(".") };
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    let idx = self.import(&module, &func);
                    self.emit_host_call(idx, arg_exprs.len() as u8);
                    return Ok(());
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
                    for a in &arg_exprs { self.compile_expr(a)?; }
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

                let full_canon = self.canon(&class_path);
                if self.defined_classes.contains(&full_canon)
                    && self.scope().resolve(head_name).is_none()
                    && self.scope().resolve_ci(head_name).is_none()
                    && self.lookup_var_type_hint(head_name).is_none()
                {
                    static_class_canon = Some(full_canon);
                }

                if static_class_canon.is_none() && class_parts.len() > 1 {
                    let short_name = class_parts.last().map(String::as_str).unwrap_or("");
                    let short_canon = self.canon(short_name);
                    if self.defined_classes.contains(&short_canon)
                        && self.scope().resolve(short_name).is_none()
                        && self.scope().resolve_ci(short_name).is_none()
                        && self.lookup_var_type_hint(short_name).is_none()
                    {
                        static_class_canon = Some(short_canon);
                    }
                }

                if let Some(canon) = static_class_canon {
                    if self.is_js_profile() {
                        let cls_idx = self.str_const(&canon);
                        self.emit_u16(Op::GLOBAL_GET, cls_idx);
                        let cls_tmp = self.scope().resolve("__static_cls")
                            .unwrap_or_else(|| self.define_local("__static_cls"));
                        self.emit_u16(Op::LOCAL_SET, cls_tmp); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, cls_tmp);
                        let method_idx = self.str_const(&self.canon(field));
                        self.emit_u16(Op::STRUCT_GET, method_idx);
                        let fn_tmp = self.scope().resolve("__static_fn")
                            .unwrap_or_else(|| self.define_local("__static_fn"));
                        self.emit_u16(Op::LOCAL_SET, fn_tmp); self.emit(Op::DROP);
                        let saved_js_this = self.save_js_this("__js_prev_this_static_method");
                        self.emit_u16(Op::LOCAL_GET, cls_tmp);
                        self.set_js_this_from_stack();
                        let method_name = self.canon(field);
                        let qualified_method = self.canon(&format!("{}.{}", canon, field));
                        if let Some(param_modes) = self
                            .function_param_modes
                            .get(&qualified_method)
                            .cloned()
                            .or_else(|| self.function_param_modes.get(&method_name).cloned())
                        {
                            if param_modes.iter().any(|mode| matches!(mode, PassBy::Ref | PassBy::Out)) {
                                let mut arg_slots = Vec::with_capacity(args.len());
                                for (index, arg) in args.iter().enumerate() {
                                    match param_modes.get(index).copied().unwrap_or(PassBy::Value) {
                                        PassBy::Out => self.emit(Op::NULL),
                                        PassBy::Ref | PassBy::Const | PassBy::Value => {
                                            if !matches!(param_modes.get(index), Some(PassBy::Out)) {
                                                self.compile_expr_with_value_copy(&arg.value)?;
                                            }
                                        }
                                    }
                                    let arg_slot = self.define_local(&format!("__js_static_call_arg_{}", index));
                                    self.emit_u16(Op::LOCAL_SET, arg_slot);
                                    self.emit(Op::DROP);
                                    arg_slots.push(arg_slot);
                                }

                                self.emit_u16(Op::LOCAL_GET, fn_tmp);
                                for slot in &arg_slots {
                                    self.emit_u16(Op::LOCAL_GET, *slot);
                                }
                                self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                                let pack_slot = self.define_local("__js_static_ref_call_pack");
                                self.emit_u16(Op::LOCAL_SET, pack_slot);
                                self.emit(Op::DROP);
                                self.restore_js_this(saved_js_this);

                                let mut ref_out_index = 1usize;
                                for (index, arg) in args.iter().enumerate() {
                                    if !matches!(param_modes.get(index), Some(PassBy::Ref | PassBy::Out)) {
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

                        self.emit_u16(Op::LOCAL_GET, fn_tmp);
                        for a in &arg_exprs { self.compile_expr(a)?; }
                        self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                        let result_slot = self.define_local("__js_static_method_result");
                        self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                        self.restore_js_this(saved_js_this);
                        if args.iter().any(|arg| arg.by_ref) {
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
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                        return Ok(());
                    }

                    let cls_idx = self.str_const(&canon);
                    self.emit_u16(Op::GLOBAL_GET, cls_idx);
                    self.emit(Op::DUP);
                    let m = self.canon(field);
                    let method_idx = self.str_const(&m);
                    self.emit_u16(Op::STRUCT_GET, method_idx);
                    // Stack: [class, fn] — swap so we have [fn, class, ...args]
                    let fn_tmp = self.scope().resolve("__static_fn")
                        .unwrap_or_else(|| self.define_local("__static_fn"));
                    self.emit_u16(Op::LOCAL_SET, fn_tmp); self.emit(Op::DROP);
                    let cls_tmp = self.scope().resolve("__static_cls")
                        .unwrap_or_else(|| self.define_local("__static_cls"));
                    self.emit_u16(Op::LOCAL_SET, cls_tmp); self.emit(Op::DROP);
                    let qualified_method = self.canon(&format!("{}.{}", canon, field));
                    if let Some(param_modes) = self
                        .function_param_modes
                        .get(&qualified_method)
                        .cloned()
                        .or_else(|| self.function_param_modes.get(&m).cloned())
                    {
                        if param_modes.iter().any(|mode| matches!(mode, PassBy::Ref | PassBy::Out)) {
                            let mut arg_slots = Vec::with_capacity(args.len());
                            for (index, arg) in args.iter().enumerate() {
                                match param_modes.get(index).copied().unwrap_or(PassBy::Value) {
                                    PassBy::Out => self.emit(Op::NULL),
                                    PassBy::Ref | PassBy::Const | PassBy::Value => {
                                        if !matches!(param_modes.get(index), Some(PassBy::Out)) {
                                            self.compile_expr_with_value_copy(&arg.value)?;
                                        }
                                    }
                                }
                                let arg_slot = self.define_local(&format!("__static_call_arg_{}", index));
                                self.emit_u16(Op::LOCAL_SET, arg_slot);
                                self.emit(Op::DROP);
                                arg_slots.push(arg_slot);
                            }

                            self.emit_u16(Op::LOCAL_GET, fn_tmp);
                            for slot in &arg_slots {
                                self.emit_u16(Op::LOCAL_GET, *slot);
                            }
                            self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);

                            let pack_slot = self.define_local("__static_ref_call_pack");
                            self.emit_u16(Op::LOCAL_SET, pack_slot);
                            self.emit(Op::DROP);
                            let mut ref_out_index = 1usize;
                            for (index, arg) in args.iter().enumerate() {
                                if !matches!(param_modes.get(index), Some(PassBy::Ref | PassBy::Out)) {
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

                    if self.profile.name == "csharp" && args.len() == 1 && !args[0].spread {
                        if self.resolve_static_method_overload_for_type(&canon, field, &arg_exprs)
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
                            let arg_slot = self.define_local(&format!("__static_class_arg_{}", index));
                            self.emit_u16(Op::LOCAL_SET, arg_slot);
                            self.emit(Op::DROP);
                            arg_slots.push(arg_slot);
                        }
                        self.emit_call_ref_with_arg_slots(fn_tmp, None, &arg_slots);
                    }
                    if args.iter().any(|arg| arg.by_ref) {
                        let pack_slot = self.define_local("__static_by_ref_pack");
                        self.emit_u16(Op::LOCAL_SET, pack_slot);
                        self.emit(Op::DROP);
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
        }

        // ── Nested static type call: Outer.Inner.Method(args) ───────
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            if let ExprKind::Member { object: outer_obj, field: nested_name, .. } = &object.kind {
                if let ExprKind::Ident(outer_name) = &outer_obj.kind {
                    let outer_canon = self.canon(outer_name);
                    let is_outer_class = self.defined_classes.contains(&outer_canon)
                        && self.scope().resolve(outer_name).is_none();
                    if is_outer_class {
                        let nested_ok = self.pending_classes.get(outer_canon.as_str())
                            .map(|pc| pc.nested_types.iter().any(|n| {
                                if self.case_sensitive { n == nested_name } else { n.eq_ignore_ascii_case(nested_name) }
                            }))
                            .unwrap_or(false);
                        if nested_ok {
                            let outer_idx = self.str_const(&outer_canon);
                            self.emit_u16(Op::GLOBAL_GET, outer_idx);
                            let nested_idx = self.str_const(&self.canon(nested_name));
                            self.emit_u16(Op::STRUCT_GET, nested_idx);
                            let cls_tmp = self.scope().resolve("__nested_static_cls")
                                .unwrap_or_else(|| self.define_local("__nested_static_cls"));
                            self.emit_u16(Op::LOCAL_SET, cls_tmp); self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_GET, cls_tmp);
                            let method_idx = self.str_const(&self.canon(field));
                            self.emit_u16(Op::STRUCT_GET, method_idx);
                            let fn_tmp = self.scope().resolve("__nested_static_fn")
                                .unwrap_or_else(|| self.define_local("__nested_static_fn"));
                            self.emit_u16(Op::LOCAL_SET, fn_tmp); self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_GET, fn_tmp);
                            for a in &arg_exprs { self.compile_expr(a)?; }
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
            if !self.direct_receiver_has_own_pending_method(object, field)
                && (field == "call" || field == "apply")
            {
                let saved_js_this = self.save_js_this("__js_prev_this_call");
                if self.is_js_profile() {
                    if let Some(this_arg) = arg_exprs.first() {
                        self.compile_expr(this_arg)?;
                    } else {
                        let line = self.line;
                        common::expressions::emit_undefined(self.chunk(), line);
                    }
                    self.set_js_this_from_stack();
                }
                self.compile_expr(object)?;                       // [fn]
                if field == "call" {
                    // Skip thisArg, compile rest as positional args.
                    for a in arg_exprs.iter().skip(1) {
                        self.compile_expr(a)?;
                    }
                    let n = arg_exprs.len().saturating_sub(1);
                    self.emit_u8(Op::CALL_REF, n as u8);
                } else {
                    // apply(thisArg, argsArray) — spread the array.
                    if let Some(args_expr) = arg_exprs.get(1) {
                        self.compile_expr(args_expr)?;
                        self.emit(Op::SPREAD);
                    }
                    // Use call_ref with 0 — the spread opcode pushes
                    // each array element and bumps the call arity at
                    // runtime via Op::call_spread if available, else
                    // we fall back here. The current VM uses Op::SPREAD
                    // before call_ref to flatten the top array.
                    self.emit_u8(Op::CALL_REF, 0);
                }
                if saved_js_this.is_some() {
                    let result_slot = self.define_local("__js_call_result");
                    self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                    self.restore_js_this(saved_js_this);
                    self.emit_u16(Op::LOCAL_GET, result_slot);
                }
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
                let class_name = Self::normalize_type_hint(&class_name);
                let surface = common::dotnet::surface();
                if let Some(target) = surface.lookup_instance_method(&class_name, field, arg_exprs.len() as u8) {
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
                                then: Box::new(Expression::new(ExprKind::Lit(Literal::Int(-1)))),
                                else_: Box::new(Expression::new(ExprKind::Ternary {
                                    cond: Box::new(Expression::new(ExprKind::Binary {
                                        op: BinOp::Gt,
                                        left: Box::new(Expression::ident("left")),
                                        right: Box::new(Expression::ident("right")),
                                    })),
                                    then: Box::new(Expression::new(ExprKind::Lit(Literal::Int(1)))),
                                    else_: Box::new(Expression::new(ExprKind::Lit(Literal::Int(0)))),
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
                        && class_name.rsplit('.').next().is_some_and(|name| name.eq_ignore_ascii_case("List") || name.eq_ignore_ascii_case("ArrayList"))
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
                    for a in &arg_exprs { self.compile_expr(a)?; }
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
        if let ExprKind::Member { object, field, null_safe } = &callee.kind {
            let canon_field = self.canon(field);
            let receiver_is_direct = matches!(
                object.kind,
                ExprKind::This | ExprKind::Super | ExprKind::Ident(_)
            );
            if self.is_python_profile() && arg_exprs.is_empty() {
                if let ExprKind::Lit(Literal::Str(value)) = &object.kind {
                    match field.as_str() {
                        "isidentifier" => {
                            self.emit_const(Value::Bool(python_is_identifier_literal(value.as_ref())));
                            return Ok(());
                        }
                        "isprintable" => {
                            self.emit_const(Value::Bool(python_is_printable_literal(value.as_ref())));
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
                self.profile.lookup_value_method(field, arg_exprs.len() as u8).cloned()
            };
            let prefer_string_stdlib_value_method = matches!(
                matched_value_method.as_ref().map(|d| &d.emit),
                Some(BuiltinEmit::Stdlib(_))
            ) && self.expr_is_known_string_receiver(object);
            // Keep dotnet adapter value-methods ahead of runtime collection
            // dispatch for untyped receivers (notably plain arrays using
            // LINQ-style extension methods like Select/SelectMany).
            let prefer_dotnet_adapter = match matched_value_method.as_ref().map(|d| &d.emit) {
                Some(BuiltinEmit::Common(name)) => name.starts_with("dotnet."),
                _ => false,
            };
            let receiver_has_pending_user_method = self
                .infer_expr_type_hint(object)
                .as_deref()
                .is_some_and(|type_hint| self.pending_class_has_method_name_for_type(type_hint, field));
            let receiver_is_user_type = self
                .infer_expr_type_hint(object)
                .as_deref()
                .and_then(|type_hint| self.resolve_pending_class_name_for_type_hint(type_hint))
                .is_some();
            let user_method_shadow = self.direct_receiver_has_own_pending_method(object, field)
                || receiver_has_pending_user_method
                || (receiver_is_direct && receiver_is_user_type && self.defined_class_methods.contains(&canon_field));
            // Also skip value_methods if the field is an array HOF method —
            // the array_methods dispatch handles it with proper HOF semantics.
            // Without this, `[1,2,3].includes(2)` routes through the string
            // `includes` value method instead of the array contains HOF.
            let field_lower_check = if self.case_sensitive { field.clone() } else { field.to_lowercase() };
            let is_array_method = self.profile.lookup_array_method(&field_lower_check).is_some();
            if user_method_shadow || is_array_method {
                // Fall through — let the HOF dispatch or generic call path handle it
            } else if self.profile.namespaces.use_dotnet
                && common::dotnet::uses_runtime_collection_dispatch_arity(field, arg_exprs.len() as u8)
                && !prefer_string_stdlib_value_method
                && !prefer_dotnet_adapter
            {
                // Let the generic member-call path consult the runtime type
                // registry for shared .NET collection methods instead of
                // intercepting them via language profile value-method tables.
            } else if let Some(def) = matched_value_method {
                // For Stdlib calls, push func ref BEFORE args (call_ref expects [func, args...])
                if let BuiltinEmit::Stdlib(stdlib_name) = &def.emit {
                    let global_name = Self::stdlib_global_name(stdlib_name);
                    let name_idx = self.str_const(&global_name);
                    self.emit_u16(Op::GLOBAL_GET, name_idx);
                    self.compile_expr(object)?;
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                    return Ok(());
                }
                // Object is first arg, then explicit args
                self.compile_expr(object)?;
                for a in &arg_exprs { self.compile_expr(a)?; }
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
                            self.emit(Op::I32_CONST_0); // start
                            self.emit_const(Value::I32(i32::MAX)); // end (clamped by VM)
                        }
                        // C# `s.Substring(start)` — 1-arg form means
                        // "from start to end of string". STR_SUBSTRING
                        // wants `[s, start, end]`; default end to a
                        // sentinel large value (VM clamps to s.len()).
                        // Same shape applies to ECMA-262 §22.1.3.16
                        // `String.prototype.slice(start)`.
                        "strings.substring" | "strings.slice"
                            if arg_exprs.len() < 2 => {
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
            let field_lower = if self.case_sensitive { field.clone() } else { field.to_lowercase() };
            let user_class_method = self.direct_receiver_has_own_pending_method(object, field)
                || self
                    .infer_expr_type_hint(object)
                    .as_deref()
                    .is_some_and(|type_hint| self.pending_class_has_method_name_for_type(type_hint, field));
            if !user_class_method
                && self.profile.lookup_array_method(&field_lower).is_some()
            {
                // (re-fetch only when we're committed to the HOF path so
                // the method name lookup matches the previous behaviour)
            }
            if let Some(stdlib_name) = self.profile.lookup_array_method(&field_lower)
                .filter(|_| !user_class_method)
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
                self.emit_u16(Op::LOCAL_SET, arr_slot); self.emit(Op::DROP);

                if let Some(fn_expr) = arg_exprs.first() {
                    self.compile_expr(fn_expr)?;
                } else {
                    self.emit(Op::NULL);
                }
                let fn_slot = self.define_local("__hof_fn");
                self.emit_u16(Op::LOCAL_SET, fn_slot); self.emit(Op::DROP);

                let idx_slot = self.define_local("__hof_idx");
                let result_slot = self.define_local("__hof_result");
                let line = self.line;

                match field_lower.as_str() {
                    "map" => {
                        // emit_map leaves result on stack
                        common::loops::emit_map(&mut self.chunks, self.current, fn_slot, arr_slot, result_slot, idx_slot, line);
                    }
                    "filter" => {
                        let elem_slot = self.define_local("__hof_elem");
                        common::loops::emit_filter(&mut self.chunks, self.current, fn_slot, arr_slot, result_slot, idx_slot, elem_slot, line);
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
                            self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                            // Inline reduce loop starting from i=0
                            self.emit(Op::I32_CONST_0);
                            self.emit_u16(Op::LOCAL_SET, idx_slot); self.emit(Op::DROP);
                            let loop_start = self.chunks[self.current].current_offset();
                            self.emit_u16(Op::LOCAL_GET, idx_slot);
                            self.emit_u16(Op::LOCAL_GET, arr_slot);
                            { let l = self.line; common::collections::emit_len(&mut self.chunks, self.current, l); }
                            self.emit(Op::DYN_LT);
                            let exit_jump = self.emit_jump(Op::BR_IF_FALSE);
                            // acc = fn(acc, arr[i])
                            self.emit_u16(Op::LOCAL_GET, fn_slot);
                            self.emit_u16(Op::LOCAL_GET, result_slot);
                            self.emit_u16(Op::LOCAL_GET, arr_slot);
                            self.emit_u16(Op::LOCAL_GET, idx_slot);
                            { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
                            self.emit_u8(Op::CALL_REF, 2);
                            self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                            // i++
                            self.emit_u16(Op::LOCAL_GET, idx_slot);
                            self.emit_const(Value::I32(1));
                            self.emit(Op::DYN_ADD);
                            self.emit_u16(Op::LOCAL_SET, idx_slot); self.emit(Op::DROP);
                            self.emit_loop(loop_start);
                            self.patch_jump(exit_jump);
                            self.emit_u16(Op::LOCAL_GET, result_slot);
                        } else {
                            // No initial: emit_reduce starts from arr[0], i=1
                            common::loops::emit_reduce(&mut self.chunks, self.current, fn_slot, arr_slot, result_slot, idx_slot, line);
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
                            common::invoke::emit_invoke_method(
                                &mut self.chunks,
                                self.current,
                                "forEach",
                                1,
                                line,
                            );
                            self.emit(Op::DROP); // forEach returns undefined
                        } else {
                            common::loops::emit_foreach(&mut self.chunks, self.current, fn_slot, arr_slot, idx_slot, line);
                        }
                    }
                    "some" => {
                        common::loops::emit_any_every(&mut self.chunks, self.current, fn_slot, arr_slot, idx_slot, true, line);
                    }
                    "every" => {
                        common::loops::emit_any_every(&mut self.chunks, self.current, fn_slot, arr_slot, idx_slot, false, line);
                    }
                    "find" => {
                        // find uses includes pattern but returns element not bool.
                        // JS spec §23.1.3.10: returns undefined when no match;
                        // other languages stick with Null for cross-compat
                        // (Python None / VB Nothing / .NET null match Null).
                        if self.is_js_profile() {
                            self.emit(Op::UNDEFINED);
                        } else {
                            self.emit(Op::NULL);
                        }
                        self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                        let lp = common::loops::emit_for_in_start(
                            &mut self.chunks, self.current, arr_slot, idx_slot, line);
                        let elem_slot = self.define_local("__find_elem");
                        self.emit_u16(Op::LOCAL_SET, elem_slot); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u16(Op::LOCAL_GET, elem_slot);
                        self.emit_u8(Op::CALL_REF, 1);
                        self.emit(Op::DYN_TO_BOOL);
                        let skip = self.emit_jump(Op::BR_IF_FALSE);
                        self.emit_u16(Op::LOCAL_GET, elem_slot);
                        self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                        let brk = self.emit_jump(Op::BR);
                        self.patch_jump(skip);
                        common::loops::emit_for_in_end(&mut self.chunks, self.current, idx_slot, lp, line);
                        self.patch_jump(brk);
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                    }
                    "findIndex" | "findindex" => {
                        // findIndex: like find but returns the index, not the element
                        self.emit_const(Value::I32(-1));
                        self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                        let lp = common::loops::emit_for_in_start(
                            &mut self.chunks, self.current, arr_slot, idx_slot, line);
                        let elem_slot = self.define_local("__findi_elem");
                        self.emit_u16(Op::LOCAL_SET, elem_slot); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u16(Op::LOCAL_GET, elem_slot);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_u8(Op::CALL_REF, 2);
                        self.emit(Op::DYN_TO_BOOL);
                        let skip = self.emit_jump(Op::BR_IF_FALSE);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                        let brk = self.emit_jump(Op::BR);
                        self.patch_jump(skip);
                        common::loops::emit_for_in_end(&mut self.chunks, self.current, idx_slot, lp, line);
                        self.patch_jump(brk);
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
                        let no_fn = self.emit_jump(Op::BR_IF_TRUE);
                        let global = self.str_const("__vybe_sort_with_comparator");
                        self.emit_u16(Op::GLOBAL_GET, global);
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u8(Op::CALL_REF, 2);
                        let done = self.emit_jump(Op::BR);
                        self.patch_jump(no_fn);
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
                                    then: Box::new(Expression::new(ExprKind::Lit(Literal::Int(-1)))),
                                    else_: Box::new(Expression::new(ExprKind::Ternary {
                                        cond: Box::new(Expression::new(ExprKind::Binary {
                                            op: BinOp::Gt,
                                            left: Box::new(Expression::ident("left")),
                                            right: Box::new(Expression::ident("right")),
                                        })),
                                        then: Box::new(Expression::new(ExprKind::Lit(Literal::Int(1)))),
                                        else_: Box::new(Expression::new(ExprKind::Lit(Literal::Int(0)))),
                                    })),
                                }))),
                                &[],
                            )?;
                            self.emit_u8(Op::CALL_REF, 2);
                        }
                        self.patch_jump(done);
                    }
                    "sort_by_key" => {
                        // .NET OrderBy(keySelector) — 1-arg key extractor
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit(Op::REF_IS_NULL);
                        let no_fn = self.emit_jump(Op::BR_IF_TRUE);
                        let global = self.str_const("__vybe_sort_by_key");
                        self.emit_u16(Op::GLOBAL_GET, global);
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u8(Op::CALL_REF, 2);
                        let done = self.emit_jump(Op::BR);
                        self.patch_jump(no_fn);
                        let sort_global = self.str_const("__vybe_sort_in_place");
                        self.emit_u16(Op::GLOBAL_GET, sort_global);
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u8(Op::CALL_REF, 1);
                        self.patch_jump(done);
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
                        common::loops::emit_map(&mut self.chunks, self.current, fn_slot, arr_slot, mapped_slot, idx_slot, line);
                        // Now the mapped array is on stack. Flatten it one level.
                        let flat_idx = self.import("ecma:array", "flat");
                        self.emit_const(Value::I32(1));  // depth = 1
                        self.emit_host_call(flat_idx, 2);
                    }
                    "reduceRight" | "reduceright" => {
                        // reduceRight(fn, initial?) — iterate from end to start.
                        if let Some(init_expr) = arg_exprs.get(1) {
                            self.compile_expr(init_expr)?;
                            self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                        } else {
                            // acc = arr[len-1]
                            self.emit_u16(Op::LOCAL_GET, arr_slot);
                            { let l = self.line; common::collections::emit_len(&mut self.chunks, self.current, l); }
                            self.emit_const(Value::I32(1));
                            self.emit(Op::F64_SUB);
                            self.emit_u16(Op::LOCAL_SET, idx_slot); self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_GET, arr_slot);
                            self.emit_u16(Op::LOCAL_GET, idx_slot);
                            { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
                            self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                        }
                        // Start from len-1 (or len-2 if no initial)
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        { let l = self.line; common::collections::emit_len(&mut self.chunks, self.current, l); }
                        self.emit_const(Value::I32(1));
                        self.emit(Op::F64_SUB);
                        if arg_exprs.get(1).is_none() {
                            self.emit_const(Value::I32(1));
                            self.emit(Op::F64_SUB);
                        }
                        self.emit_u16(Op::LOCAL_SET, idx_slot); self.emit(Op::DROP);
                        let loop_start = self.chunks[self.current].current_offset();
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_const(Value::I32(0));
                        self.emit(Op::DYN_GE);
                        let exit_jump = self.emit_jump(Op::BR_IF_FALSE);
                        // acc = fn(acc, arr[i])
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
                        self.emit_u8(Op::CALL_REF, 2);
                        self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                        // i--
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_const(Value::I32(1));
                        self.emit(Op::F64_SUB);
                        self.emit_u16(Op::LOCAL_SET, idx_slot); self.emit(Op::DROP);
                        self.emit_loop(loop_start);
                        self.patch_jump(exit_jump);
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                    }
                    "findLast" | "findlast" => {
                        // Iterate backward, return last element matching predicate
                        self.emit(Op::NULL);
                        self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        { let l = self.line; common::collections::emit_len(&mut self.chunks, self.current, l); }
                        self.emit_const(Value::I32(1));
                        self.emit(Op::F64_SUB);
                        self.emit_u16(Op::LOCAL_SET, idx_slot); self.emit(Op::DROP);
                        let loop_start = self.chunks[self.current].current_offset();
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_const(Value::I32(0));
                        self.emit(Op::DYN_GE);
                        let exit_jump = self.emit_jump(Op::BR_IF_FALSE);
                        let elem_slot = self.define_local("__fl_elem");
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
                        self.emit_u16(Op::LOCAL_SET, elem_slot); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u16(Op::LOCAL_GET, elem_slot);
                        self.emit_u8(Op::CALL_REF, 1);
                        self.emit(Op::DYN_TO_BOOL);
                        let skip = self.emit_jump(Op::BR_IF_FALSE);
                        self.emit_u16(Op::LOCAL_GET, elem_slot);
                        self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                        let brk = self.emit_jump(Op::BR);
                        self.patch_jump(skip);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_const(Value::I32(1));
                        self.emit(Op::F64_SUB);
                        self.emit_u16(Op::LOCAL_SET, idx_slot); self.emit(Op::DROP);
                        self.emit_loop(loop_start);
                        self.patch_jump(exit_jump);
                        self.patch_jump(brk);
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                    }
                    "findLastIndex" | "findlastindex" => {
                        // Iterate backward, return last index matching predicate
                        self.emit_const(Value::I32(-1));
                        self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        { let l = self.line; common::collections::emit_len(&mut self.chunks, self.current, l); }
                        self.emit_const(Value::I32(1));
                        self.emit(Op::F64_SUB);
                        self.emit_u16(Op::LOCAL_SET, idx_slot); self.emit(Op::DROP);
                        let loop_start = self.chunks[self.current].current_offset();
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_const(Value::I32(0));
                        self.emit(Op::DYN_GE);
                        let exit_jump = self.emit_jump(Op::BR_IF_FALSE);
                        let elem_slot2 = self.define_local("__fli_elem");
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
                        self.emit_u16(Op::LOCAL_SET, elem_slot2); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u16(Op::LOCAL_GET, elem_slot2);
                        self.emit_u8(Op::CALL_REF, 1);
                        self.emit(Op::DYN_TO_BOOL);
                        let skip = self.emit_jump(Op::BR_IF_FALSE);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                        let brk = self.emit_jump(Op::BR);
                        self.patch_jump(skip);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_const(Value::I32(1));
                        self.emit(Op::F64_SUB);
                        self.emit_u16(Op::LOCAL_SET, idx_slot); self.emit(Op::DROP);
                        self.emit_loop(loop_start);
                        self.patch_jump(exit_jump);
                        self.patch_jump(brk);
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                    }
                    "removeAll" | "removeall" => {
                        // Iterate backward over arr, splice each matching element.
                        // Returns count of removed items.
                        let removed_slot = self.define_local("__ra_removed");
                        self.emit_const(Value::I32(0));
                        self.emit_u16(Op::LOCAL_SET, removed_slot); self.emit(Op::DROP);
                        // Start i = arr.len - 1
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        { let l = self.line; common::collections::emit_len(&mut self.chunks, self.current, l); }
                        self.emit_const(Value::I32(1));
                        self.emit(Op::F64_SUB);
                        self.emit_u16(Op::LOCAL_SET, idx_slot); self.emit(Op::DROP);
                        let ra_loop = self.chunks[self.current].current_offset();
                        // while i >= 0
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_const(Value::I32(0));
                        self.emit(Op::DYN_GE);
                        let ra_exit = self.emit_jump(Op::BR_IF_FALSE);
                        // elem = arr[i]
                        let ra_elem = self.define_local("__ra_elem");
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
                        self.emit_u16(Op::LOCAL_SET, ra_elem); self.emit(Op::DROP);
                        // if fn(elem) → remove
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u16(Op::LOCAL_GET, ra_elem);
                        self.emit_u8(Op::CALL_REF, 1);
                        self.emit(Op::DYN_TO_BOOL);
                        let ra_skip = self.emit_jump(Op::BR_IF_FALSE);
                        // splice(arr, i, 1) → drop removed array
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        { let l = self.line; common::collections::emit_remove_at(&mut self.chunks, self.current, l); }
                        self.emit(Op::DROP);
                        // removed++
                        self.emit_u16(Op::LOCAL_GET, removed_slot);
                        self.emit_const(Value::I32(1));
                        self.emit(Op::DYN_ADD);
                        self.emit_u16(Op::LOCAL_SET, removed_slot); self.emit(Op::DROP);
                        self.patch_jump(ra_skip);
                        // i--
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_const(Value::I32(1));
                        self.emit(Op::F64_SUB);
                        self.emit_u16(Op::LOCAL_SET, idx_slot); self.emit(Op::DROP);
                        self.emit_loop(ra_loop);
                        self.patch_jump(ra_exit);
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
                let is_ctor = if self.case_sensitive { field == ctor_nm } else { field.eq_ignore_ascii_case(ctor_nm) };
                let canon_class = self.canon(class_name);
                let is_known_class = self.defined_classes.contains(&canon_class)
                    && self.scope().resolve(class_name).is_none();
                if is_ctor && is_known_class {
                    self.emit_var_get(class_name);
                    for a in &arg_exprs { self.compile_expr(a)?; }
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
                        for a in &arg_exprs { self.compile_expr(a)?; }
                        self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                        return Ok(());
                    }

                    let canon_type = self.canon(&type_name);
                    let canon_field = self.canon(field);
                    let is_callable_field = self.pending_classes.get(canon_type.as_str())
                        .map(|pc| pc.fields.iter().any(|name| name == &canon_field))
                        .unwrap_or(false);
                    if is_callable_field {
                        self.compile_expr(object)?;
                        let obj_tmp = self.define_local("__pascal_callable_field_obj");
                        self.emit_u16(Op::LOCAL_SET, obj_tmp); self.emit(Op::DROP);
                        let prop = self.str_const(&canon_field);
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.emit_u16(Op::STRUCT_GET, prop);
                        for a in &arg_exprs { self.compile_expr(a)?; }
                        self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                        return Ok(());
                    }
                }
            }
        }

        // ── Method call: obj.method(args) ───────────────────────────
        if let ExprKind::Member { object, field, null_safe } = &callee.kind {
            if self.is_js_profile() {
                self.compile_expr(object)?;
                let obj_tmp = self.define_local("__js_obj");
                self.emit_u16(Op::LOCAL_SET, obj_tmp); self.emit(Op::DROP);

                let method_name = self.canon(field);
                let prop = self.str_const(&method_name);
                let receiver_marker = self.str_const("__vybe_method_receiver");

                // Generator `.return(v)`: drive the shared generator
                // return-control packet through RESUME so suspended
                // `finally` blocks execute before the completion record
                // is materialized.
                let gen_return_skip_patch = if !*null_safe && method_name == "return" && arg_exprs.len() <= 1 {
                    let line = self.line;
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    let is_gen_idx = self.import("ecma:value", "isGenerator");
                    self.emit_host_call(is_gen_idx, 1);
                    let not_gen = self.emit_jump(Op::BR_IF_FALSE);

                    let value_slot = self.define_local("__gen_return_value");
                    let done_slot = self.define_local("__gen_return_done");
                    let returned_key = self.str_const("__vybe_gen_returned");

                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    let is_done_idx = self.import("ecma:value", "isGeneratorDone");
                    self.emit_host_call(is_done_idx, 1);
                    let not_done = self.emit_jump(Op::BR_IF_FALSE);

                    if arg_exprs.is_empty() {
                        self.emit(Op::UNDEFINED);
                    } else {
                        self.compile_expr(&arg_exprs[0])?;
                    }
                    self.emit_u16(Op::LOCAL_SET, value_slot); self.emit(Op::DROP);
                    self.emit(Op::TRUE);
                    self.emit_u16(Op::LOCAL_SET, done_slot); self.emit(Op::DROP);
                    let after_resume = self.emit_jump(Op::BR);

                    self.patch_jump(not_done);
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    if arg_exprs.is_empty() {
                        self.emit(Op::UNDEFINED);
                    } else {
                        self.compile_expr(&arg_exprs[0])?;
                    }
                    self.emit_generator_control_packet_from_stack("return");
                    self.emit_u16(Op::RESUME, 0);
                    self.emit_u16(Op::LOCAL_SET, value_slot); self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    let is_done_idx = self.import("ecma:value", "isGeneratorDone");
                    self.emit_host_call(is_done_idx, 1);
                    self.emit_u16(Op::LOCAL_SET, done_slot); self.emit(Op::DROP);

                    self.patch_jump(after_resume);
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit(Op::TRUE);
                    self.emit_u16(Op::STRUCT_SET, returned_key);
                    self.emit(Op::DROP);

                    common::dict::emit_new(&mut self.chunks, self.current, line);
                    self.emit(Op::DUP);
                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    let value_key = self.str_const("value");
                    self.emit_u16(Op::STRUCT_SET, value_key);
                    self.emit(Op::DROP);
                    self.emit(Op::DUP);
                    self.emit_u16(Op::LOCAL_GET, done_slot);
                    let done_key = self.str_const("done");
                    self.emit_u16(Op::STRUCT_SET, done_key);
                    self.emit(Op::DROP);
                    let skip = self.emit_jump(Op::BR);
                    self.patch_jump(not_gen);
                    Some(skip)
                } else { None };

                // Generator `.next()` / `.next(v)`: if receiver is a
                // Continuation, drive via WASM stack-switching opcodes
                // and wrap into spec `{value, done}`.
                //   - `g.next()`     → Op::GEN_NEXT (pushes value+has_more)
                //   - `g.next(v)`    → Op::RESUME with v as resume_val
                //                       (pushes yielded value), then
                //                       check `isGeneratorDone` for the
                //                       done flag.
                // Non-Continuations (Array iterators, custom iterables)
                // fall through to regular method dispatch below.
                let gen_next_skip_patch = if !*null_safe && method_name == "next" && arg_exprs.len() <= 1 {
                    let line = self.line;
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    let is_gen_idx = self.import("ecma:value", "isGenerator");
                    self.emit_host_call(is_gen_idx, 1);
                    let not_gen = self.emit_jump(Op::BR_IF_FALSE);
                    let value_slot = self.define_local("__gen_value");
                    let done_slot = self.define_local("__gen_done");
                    let started_key = self.str_const("__vybe_gen_started");
                    // If a previous `.return()` stamped the cont as
                    // returned, short-circuit to `{value: undefined,
                    // done: true}` per ECMA-262 §27.5.1.2 step 2.
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    let returned_key2 = self.str_const("__vybe_gen_returned");
                    self.emit_u16(Op::STRUCT_GET, returned_key2);
                    self.emit(Op::DYN_TO_BOOL);
                    let not_returned = self.emit_jump(Op::BR_IF_FALSE);
                    self.emit(Op::UNDEFINED);
                    self.emit_u16(Op::LOCAL_SET, value_slot); self.emit(Op::DROP);
                    self.emit(Op::TRUE);
                    self.emit_u16(Op::LOCAL_SET, done_slot); self.emit(Op::DROP);
                    let after_returned_branch = self.emit_jump(Op::BR);
                    self.patch_jump(not_returned);
                    if arg_exprs.is_empty() {
                        // `g.next()` — GEN_NEXT path: pushes value+has_more.
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.emit(Op::GEN_NEXT);
                        let has_more_slot = self.define_local("__gen_has_more");
                        self.emit_u16(Op::LOCAL_SET, has_more_slot); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_SET, value_slot); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, has_more_slot);
                        self.emit(Op::DYN_TO_BOOL);
                        self.emit(Op::DYN_NOT);
                        self.emit_u16(Op::LOCAL_SET, done_slot); self.emit(Op::DROP);
                    } else {
                        // `g.next(v)` — RESUME with the resume value;
                        // the suspended yield expression evaluates to
                        // `v`. Pushes only the yielded value back; we
                        // query `isGeneratorDone` for the spec `done`.
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.compile_expr(&arg_exprs[0])?;
                        self.emit_u16(Op::RESUME, 0);
                        self.emit_u16(Op::LOCAL_SET, value_slot); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        let is_done_idx = self.import("ecma:value", "isGeneratorDone");
                        self.emit_host_call(is_done_idx, 1);
                        self.emit_u16(Op::LOCAL_SET, done_slot); self.emit(Op::DROP);
                    }
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit(Op::TRUE);
                    self.emit_u16(Op::STRUCT_SET, started_key);
                    self.emit(Op::DROP);
                    // Both the early-`returned` short-circuit and the
                    // GEN_NEXT/RESUME paths converge here to build the
                    // `{value, done}` wrapper.
                    self.patch_jump(after_returned_branch);
                    common::dict::emit_new(&mut self.chunks, self.current, line);
                    self.emit(Op::DUP);
                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    let value_key = self.str_const("value");
                    self.emit_u16(Op::STRUCT_SET, value_key);
                    self.emit(Op::DROP);
                    self.emit(Op::DUP);
                    self.emit_u16(Op::LOCAL_GET, done_slot);
                    let done_key = self.str_const("done");
                    self.emit_u16(Op::STRUCT_SET, done_key);
                    self.emit(Op::DROP);
                    let skip = self.emit_jump(Op::BR);
                    self.patch_jump(not_gen);
                    Some(skip)
                } else { None };

                let gen_throw_skip_patch = if !*null_safe && method_name == "throw" && arg_exprs.len() <= 1 {
                    let line = self.line;
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    let is_gen_idx = self.import("ecma:value", "isGenerator");
                    self.emit_host_call(is_gen_idx, 1);
                    let not_gen = self.emit_jump(Op::BR_IF_FALSE);

                    let value_slot = self.define_local("__gen_throw_value");
                    let done_slot = self.define_local("__gen_throw_done");
                    let started_key = self.str_const("__vybe_gen_started");

                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit_u16(Op::STRUCT_GET, started_key);
                    self.emit(Op::DYN_TO_BOOL);
                    let already_started = self.emit_jump(Op::BR_IF_TRUE);

                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit(Op::GEN_NEXT);
                    let has_more_slot = self.define_local("__gen_throw_has_more");
                    self.emit_u16(Op::LOCAL_SET, has_more_slot); self.emit(Op::DROP);
                    let primed_value_slot = self.define_local("__gen_throw_primed_value");
                    self.emit_u16(Op::LOCAL_SET, primed_value_slot); self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit(Op::TRUE);
                    self.emit_u16(Op::STRUCT_SET, started_key);
                    self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, has_more_slot);
                    self.emit(Op::DYN_TO_BOOL);
                    let primed = self.emit_jump(Op::BR_IF_TRUE);
                    if arg_exprs.is_empty() {
                        self.emit(Op::UNDEFINED);
                    } else {
                        self.compile_expr(&arg_exprs[0])?;
                    }
                    self.emit(Op::THROW);

                    self.patch_jump(primed);
                    self.patch_jump(already_started);

                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    if arg_exprs.is_empty() {
                        self.emit(Op::UNDEFINED);
                    } else {
                        self.compile_expr(&arg_exprs[0])?;
                    }
                    self.emit_generator_control_packet_from_stack("throw");
                    self.emit_u16(Op::RESUME, 0);
                    self.emit_u16(Op::LOCAL_SET, value_slot); self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    let is_done_idx = self.import("ecma:value", "isGeneratorDone");
                    self.emit_host_call(is_done_idx, 1);
                    self.emit_u16(Op::LOCAL_SET, done_slot); self.emit(Op::DROP);

                    common::dict::emit_new(&mut self.chunks, self.current, line);
                    self.emit(Op::DUP);
                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    let value_key = self.str_const("value");
                    self.emit_u16(Op::STRUCT_SET, value_key);
                    self.emit(Op::DROP);
                    self.emit(Op::DUP);
                    self.emit_u16(Op::LOCAL_GET, done_slot);
                    let done_key = self.str_const("done");
                    self.emit_u16(Op::STRUCT_SET, done_key);
                    self.emit(Op::DROP);
                    let skip = self.emit_jump(Op::BR);
                    self.patch_jump(not_gen);
                    Some(skip)
                } else { None };
                let _ = gen_next_skip_patch;
                let _ = gen_return_skip_patch;
                let _ = gen_throw_skip_patch;
                // gen_next_skip_patch / gen_return_skip_patch are
                // patched at the end of the JS method dispatch (when
                // result is on stack and we'd otherwise `return Ok(())`).

                if *null_safe {
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit(Op::REF_IS_NULL);
                    let skip = self.emit_jump(Op::BR_IF_TRUE);

                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit_u16(Op::STRUCT_GET, prop);
                    let fn_slot = self.define_local("__js_method_fn");
                    self.emit_u16(Op::LOCAL_SET, fn_slot); self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, fn_slot);
                    self.emit_u16(Op::STRUCT_GET, receiver_marker);
                    self.emit(Op::REF_IS_NULL);
                    let use_js_path = self.emit_jump(Op::BR_IF_TRUE);

                    self.emit_u16(Op::LOCAL_GET, fn_slot);
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                    let typed_done = self.emit_jump(Op::BR);

                    self.patch_jump(use_js_path);
                    self.emit_u16(Op::LOCAL_GET, fn_slot);
                    self.emit(Op::REF_IS_NULL);
                    let need_lookup = self.emit_jump(Op::BR_IF_TRUE);

                    let saved_js_this = self.save_js_this("__js_prev_this_method");
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.set_js_this_from_stack();
                    let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                    for (index, arg) in arg_exprs.iter().enumerate() {
                        self.compile_expr(arg)?;
                        let arg_slot = self.define_local(&format!("__js_method_arg_{}", index));
                        self.emit_u16(Op::LOCAL_SET, arg_slot); self.emit(Op::DROP);
                        arg_slots.push(arg_slot);
                    }
                    self.emit_call_ref_with_arg_slots(fn_slot, None, &arg_slots);
                    let result_slot = self.define_local("__js_method_result");
                    self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                    self.restore_js_this(saved_js_this);
                    self.emit_u16(Op::LOCAL_GET, result_slot);
                    let js_done = self.emit_jump(Op::BR);

                    self.patch_jump(need_lookup);
                    let lookup = self.import("ecma:value", "getMethodForCall");
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit_const(Value::String(Arc::from(method_name.as_str())));
                    self.emit_host_call(lookup, 2);
                    let lookup_slot = self.define_local("__js_lookup_fn");
                    self.emit_u16(Op::LOCAL_SET, lookup_slot); self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_GET, lookup_slot);
                    self.emit(Op::REF_IS_NULL);
                    let have_fn = self.emit_jump(Op::BR_IF_FALSE);
                    let invoke = self.import("ecma:value", "invokeMethod");
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit_const(Value::String(Arc::from(method_name.as_str())));
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    self.emit_host_call(invoke, (arg_exprs.len() + 2) as u8);
                    let after_call = self.emit_jump(Op::BR);
                    self.patch_jump(have_fn);
                    let saved_js_this = self.save_js_this("__js_prev_this_lookup");
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.set_js_this_from_stack();
                    let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                    for (index, arg) in arg_exprs.iter().enumerate() {
                        self.compile_expr(arg)?;
                        let arg_slot = self.define_local(&format!("__js_lookup_arg_{}", index));
                        self.emit_u16(Op::LOCAL_SET, arg_slot); self.emit(Op::DROP);
                        arg_slots.push(arg_slot);
                    }
                    self.emit_call_ref_with_arg_slots(lookup_slot, None, &arg_slots);
                    let result_slot = self.define_local("__js_lookup_result");
                    self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                    self.restore_js_this(saved_js_this);
                    self.emit_u16(Op::LOCAL_GET, result_slot);
                    self.patch_jump(after_call);
                    self.patch_jump(js_done);
                    self.patch_jump(typed_done);
                    let end = self.emit_jump(Op::BR);
                    self.patch_jump(skip);
                    self.emit(Op::NULL);
                    self.patch_jump(end);
                    return Ok(());
                }

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, prop);
                let fn_slot = self.define_local("__js_method_fn");
                self.emit_u16(Op::LOCAL_SET, fn_slot); self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, fn_slot);
                self.emit_u16(Op::STRUCT_GET, receiver_marker);
                self.emit(Op::REF_IS_NULL);
                let use_js_path = self.emit_jump(Op::BR_IF_TRUE);

                self.emit_u16(Op::LOCAL_GET, fn_slot);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                for a in &arg_exprs { self.compile_expr(a)?; }
                self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                let typed_done = self.emit_jump(Op::BR);

                self.patch_jump(use_js_path);
                self.emit_u16(Op::LOCAL_GET, fn_slot);
                self.emit(Op::REF_IS_NULL);
                let need_lookup = self.emit_jump(Op::BR_IF_TRUE);

                let saved_js_this = self.save_js_this("__js_prev_this_method");
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.set_js_this_from_stack();
                let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                for (index, arg) in arg_exprs.iter().enumerate() {
                    self.compile_expr(arg)?;
                    let arg_slot = self.define_local(&format!("__js_method_arg_{}", index));
                    self.emit_u16(Op::LOCAL_SET, arg_slot); self.emit(Op::DROP);
                    arg_slots.push(arg_slot);
                }
                self.emit_call_ref_with_arg_slots(fn_slot, None, &arg_slots);
                let result_slot = self.define_local("__js_method_result");
                self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                self.restore_js_this(saved_js_this);
                self.emit_u16(Op::LOCAL_GET, result_slot);
                let js_done = self.emit_jump(Op::BR);

                self.patch_jump(need_lookup);
                let lookup = self.import("ecma:value", "getMethodForCall");
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_const(Value::String(Arc::from(method_name.as_str())));
                self.emit_host_call(lookup, 2);
                let lookup_slot = self.define_local("__js_lookup_fn");
                self.emit_u16(Op::LOCAL_SET, lookup_slot); self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, lookup_slot);
                self.emit(Op::REF_IS_NULL);
                let have_fn = self.emit_jump(Op::BR_IF_FALSE);
                let invoke = self.import("ecma:value", "invokeMethod");
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_const(Value::String(Arc::from(method_name.as_str())));
                for a in &arg_exprs { self.compile_expr(a)?; }
                self.emit_host_call(invoke, (arg_exprs.len() + 2) as u8);
                let after_call = self.emit_jump(Op::BR);
                self.patch_jump(have_fn);
                let saved_js_this = self.save_js_this("__js_prev_this_lookup");
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.set_js_this_from_stack();
                let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                for (index, arg) in arg_exprs.iter().enumerate() {
                    self.compile_expr(arg)?;
                    let arg_slot = self.define_local(&format!("__js_lookup_arg_{}", index));
                    self.emit_u16(Op::LOCAL_SET, arg_slot); self.emit(Op::DROP);
                    arg_slots.push(arg_slot);
                }
                self.emit_call_ref_with_arg_slots(lookup_slot, None, &arg_slots);
                let result_slot = self.define_local("__js_lookup_result");
                self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                self.restore_js_this(saved_js_this);
                self.emit_u16(Op::LOCAL_GET, result_slot);
                self.patch_jump(after_call);
                self.patch_jump(js_done);
                self.patch_jump(typed_done);
                if let Some(skip) = gen_next_skip_patch {
                    self.patch_jump(skip);
                }
                if let Some(skip) = gen_return_skip_patch {
                    self.patch_jump(skip);
                }
                if let Some(skip) = gen_throw_skip_patch {
                    self.patch_jump(skip);
                }
                return Ok(());
            }

            self.compile_expr(object)?;
            let obj_tmp = self.define_local("__obj");
            self.reserve_local_slot(obj_tmp);
            self.emit_u16(Op::LOCAL_SET, obj_tmp); self.emit(Op::DROP);

            let field_name = self.canon(field);
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
                        self.compile_expr(arg)?;
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
                let obj_not_null = self.emit_jump(Op::BR_IF_FALSE);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                let end = self.emit_jump(Op::BR);
                self.patch_jump(obj_not_null);
                if field.eq_ignore_ascii_case("Invoke") {
                    // C# delegate null-conditional invocation: `d?.Invoke(args)`
                    // should call the delegate value directly when non-null.
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                    self.patch_jump(end);
                    return Ok(());
                }
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, prop);
                let fn_tmp = self.define_local("__fn");
                self.emit_u16(Op::LOCAL_SET, fn_tmp); self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, fn_tmp);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                for a in &arg_exprs { self.compile_expr(a)?; }
                self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                self.patch_jump(end);
                return Ok(());
            }

            let php_generator_end = if self.is_php_profile() {
                self.emit_php_generator_method_dispatch(obj_tmp, &field_name, &arg_exprs)?
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
                    let not_gen = self.emit_jump(Op::BR_IF_FALSE);

                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    match field_name.as_str() {
                        "send" => {
                            self.compile_expr(&arg_exprs[0])?;
                        }
                        "throw" => {
                            self.compile_expr(&arg_exprs[0])?;
                            self.emit_generator_control_packet_from_stack("throw");
                        }
                        "close" => {
                            self.emit(Op::NULL);
                            self.emit_generator_control_packet_from_stack("return");
                        }
                        _ => unreachable!(),
                    }
                    self.emit_u16(Op::RESUME, 0);
                    let end = self.emit_jump(Op::BR);
                    self.patch_jump(not_gen);
                    self.patch_jump(end);
                }
            }

            if let Some(end) = php_generator_end {
                if self.profile.namespaces.use_dotnet
                    && arg_exprs.is_empty()
                    && field.eq_ignore_ascii_case("sort")
                    && common::dotnet::uses_runtime_collection_dispatch_arity(field, 0)
                {
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    let line = self.line;
                    self.emit_common("dotnet.array_sort", 1, line);
                    let generic_done = self.emit_jump(Op::BR);
                    self.patch_jump(end);
                    self.patch_jump(generic_done);
                    return Ok(());
                }

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u16(Op::STRUCT_GET, prop);
                let fn_tmp = self.define_local("__fn");
                self.emit_u16(Op::LOCAL_SET, fn_tmp); self.emit(Op::DROP);
                if self.is_js_profile() {
                    let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                    for (index, arg) in arg_exprs.iter().enumerate() {
                        self.compile_expr(arg)?;
                        let arg_slot = self.define_local(&format!("__js_member_bound_arg_{}", index));
                        self.emit_u16(Op::LOCAL_SET, arg_slot);
                        self.emit(Op::DROP);
                        arg_slots.push(arg_slot);
                    }
                    self.emit_call_ref_with_arg_slots(fn_tmp, None, &arg_slots);
                    let generic_done = self.emit_jump(Op::BR);

                    self.patch_jump(end);
                    self.patch_jump(generic_done);
                    return Ok(());
                }
                if self.profile.name == "php" {
                    if let Some(overload) = self.resolve_instance_method_overload(object, field, &arg_exprs, false) {
                        let line = self.line;
                        self.emit_u16(Op::REF_FUNC, overload.chunk_idx as u16);
                        self.chunk().emit(0, line);
                        let direct_fn_tmp = self.define_local("__php_direct_instance_fn");
                        self.emit_u16(Op::LOCAL_SET, direct_fn_tmp);
                        self.emit(Op::DROP);
                        if overload.signature.has_rest {
                            self.emit_known_rest_call_from_local(direct_fn_tmp, Some(obj_tmp), args, &overload.signature)?;
                        } else {
                            self.emit_u16(Op::LOCAL_GET, direct_fn_tmp);
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            for a in &arg_exprs { self.compile_expr(a)?; }
                            self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                        }
                        let generic_done = self.emit_jump(Op::BR);

                        self.patch_jump(end);
                        self.patch_jump(generic_done);
                        return Ok(());
                    }
                }
                if resolves_to_static_container_method(self, object, field) {
                    if self.profile.name == "php" {
                        let class_canon = self.canon(&self.flatten_member_chain(object).join("."));
                        if let Some(overload) = self.resolve_static_method_overload_for_type(&class_canon, field, &arg_exprs) {
                            let line = self.line;
                            self.emit_u16(Op::REF_FUNC, overload.chunk_idx as u16);
                            self.chunk().emit(0, line);
                            let direct_fn_tmp = self.define_local("__php_direct_static_fn");
                            self.emit_u16(Op::LOCAL_SET, direct_fn_tmp);
                            self.emit(Op::DROP);
                            if overload.signature.has_rest {
                                self.emit_known_rest_call_from_local(direct_fn_tmp, Some(obj_tmp), args, &overload.signature)?;
                            } else {
                                self.emit_u16(Op::LOCAL_GET, direct_fn_tmp);
                                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                                for a in &arg_exprs { self.compile_expr(a)?; }
                                self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                            }
                            let generic_done = self.emit_jump(Op::BR);

                            self.patch_jump(end);
                            self.patch_jump(generic_done);
                            return Ok(());
                        }
                    }
                    if self.profile.name == "csharp" && args.len() == 1 && !args[0].spread {
                        let class_canon = self.canon(&self.flatten_member_chain(object).join("."));
                        if self.resolve_static_method_overload_for_type(&class_canon, field, &arg_exprs)
                            .is_some_and(|overload| overload.signature.has_rest)
                        {
                            self.emit_variadic_array_call_from_local(fn_tmp, &args[0].value)?;
                            let generic_done = self.emit_jump(Op::BR);

                            self.patch_jump(end);
                            self.patch_jump(generic_done);
                            return Ok(());
                        }
                    }
                    if let Some(overload) = self
                        .flatten_member_chain(object)
                        .last()
                        .and_then(|class_name| self.resolve_static_method_overload_for_type(class_name, field, &arg_exprs))
                        .filter(|overload| overload.signature.has_rest)
                    {
                        self.emit_known_rest_call_from_local(
                            fn_tmp,
                            if self.profile.name == "php" { Some(obj_tmp) } else { None },
                            args,
                            &overload.signature,
                        )?;
                        let generic_done = self.emit_jump(Op::BR);

                        self.patch_jump(end);
                        self.patch_jump(generic_done);
                        return Ok(());
                    }
                    let rest_signature = self
                        .flatten_member_chain(object)
                        .last()
                        .and_then(|class_name| self.resolve_static_method_overload_for_type(class_name, field, &arg_exprs))
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
                            if self.profile.name == "php" { Some(obj_tmp) } else { None },
                            args,
                            signature,
                        )?;
                    } else if self.is_js_profile() {
                        let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                        for (index, arg) in arg_exprs.iter().enumerate() {
                            self.compile_expr(arg)?;
                            let arg_slot = self.define_local(&format!("__js_static_member_arg_{}", index));
                            self.emit_u16(Op::LOCAL_SET, arg_slot);
                            self.emit(Op::DROP);
                            arg_slots.push(arg_slot);
                        }
                        self.emit_call_ref_with_arg_slots(fn_tmp, None, &arg_slots);
                    } else {
                        let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                        for (index, arg) in arg_exprs.iter().enumerate() {
                            self.compile_expr(arg)?;
                            let arg_slot = self.define_local(&format!("__static_member_arg_{}", index));
                            self.emit_u16(Op::LOCAL_SET, arg_slot);
                            self.emit(Op::DROP);
                            arg_slots.push(arg_slot);
                        }
                        self.emit_call_ref_with_arg_slots(
                            fn_tmp,
                            if self.profile.name == "php" { Some(obj_tmp) } else { None },
                            &arg_slots,
                        );
                    }
                    if args.iter().any(|arg| arg.by_ref) {
                        let pack_slot = self.define_local("__member_static_container_by_ref_pack");
                        self.emit_u16(Op::LOCAL_SET, pack_slot);
                        self.emit(Op::DROP);
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
                    let generic_done = self.emit_jump(Op::BR);

                    self.patch_jump(end);
                    self.patch_jump(generic_done);
                    return Ok(());
                }
                let primitive_tostring_done = if self.profile.namespaces.use_dotnet
                    && arg_exprs.is_empty()
                    && field.eq_ignore_ascii_case("ToString")
                {
                    let type_tmp = self.define_local("__dotnet_tostring_type");
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit(Op::REF_TYPEOF);
                    self.emit_u16(Op::LOCAL_SET, type_tmp);
                    self.emit(Op::DROP);

                    let mut primitive_matches = Vec::new();
                    for type_name in ["number", "i32", "i64", "string", "boolean"] {
                        self.emit_u16(Op::LOCAL_GET, type_tmp);
                        self.emit_const(Value::String(Arc::from(type_name)));
                        self.emit(Op::DYN_EQ);
                        primitive_matches.push(self.emit_jump(Op::BR_IF_TRUE));
                    }
                    let not_primitive = self.emit_jump(Op::BR);

                    for patch in primitive_matches {
                        self.patch_jump(patch);
                    }
                    let tostring_global = self.str_const("__vybe_tostring");
                    self.emit_u16(Op::GLOBAL_GET, tostring_global);
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit_u8(Op::CALL_REF, 1);
                    let done = self.emit_jump(Op::BR);

                    self.patch_jump(not_primitive);
                    Some(done)
                } else {
                    None
                };
                if self.profile.namespaces.use_dotnet
                    && arg_exprs.is_empty()
                    && field.eq_ignore_ascii_case("Count")
                {
                    self.emit_u16(Op::LOCAL_GET, fn_tmp);
                    self.emit(Op::REF_TYPEOF);
                    self.emit_const(Value::String(Arc::from("function")));
                    self.emit(Op::DYN_EQ);
                    let return_value = self.emit_jump(Op::BR_IF_FALSE);

                    self.emit_u16(Op::LOCAL_GET, fn_tmp);
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit_u8(Op::CALL_REF, 1);
                    let generic_done = self.emit_jump(Op::BR);

                    self.patch_jump(return_value);
                    self.emit_u16(Op::LOCAL_GET, fn_tmp);
                    self.patch_jump(generic_done);
                    return Ok(());
                }
                if self.is_js_profile() {
                    let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                    for (index, arg) in arg_exprs.iter().enumerate() {
                        self.compile_expr(arg)?;
                        let arg_slot = self.define_local(&format!("__js_member_fast_arg_{}", index));
                        self.emit_u16(Op::LOCAL_SET, arg_slot);
                        self.emit(Op::DROP);
                        arg_slots.push(arg_slot);
                    }
                    self.emit_call_ref_with_arg_slots(fn_tmp, None, &arg_slots);
                } else {
                    let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                    for (index, arg) in arg_exprs.iter().enumerate() {
                        self.compile_expr(arg)?;
                        let arg_slot = self.define_local(&format!("__member_fast_arg_{}", index));
                        self.emit_u16(Op::LOCAL_SET, arg_slot);
                        self.emit(Op::DROP);
                        arg_slots.push(arg_slot);
                    }
                    self.emit_call_ref_with_arg_slots(fn_tmp, Some(obj_tmp), &arg_slots);
                }
                if let Some(done) = primitive_tostring_done {
                    self.patch_jump(done);
                }
                let generic_done = self.emit_jump(Op::BR);

                self.patch_jump(end);
                self.patch_jump(generic_done);
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

            if let Some(chunk_idx) = self.resolve_instance_method_overload_chunk(object, field, &arg_exprs) {
                let overload = self
                    .resolve_instance_method_overload(object, field, &arg_exprs, true)
                    .ok_or_else(|| format!("failed to resolve method overload for {}", field))?;
                if overload.signature.has_rest {
                    let line = self.line;
                    self.emit_u16(Op::REF_FUNC, chunk_idx as u16);
                    self.chunk().emit(0, line);
                    let direct_fn_tmp = self.define_local("__direct_instance_fn");
                    self.emit_u16(Op::LOCAL_SET, direct_fn_tmp);
                    self.emit(Op::DROP);
                    self.emit_known_rest_call_from_local(
                        direct_fn_tmp,
                        if self.is_js_profile() { None } else { Some(obj_tmp) },
                        args,
                        &overload.signature,
                    )?;
                } else {
                    self.emit_direct_instance_method_call(chunk_idx, obj_tmp, &arg_exprs)?;
                }
                return Ok(());
            }

            if let Some(class_name) = resolve_go_pending_instance_method_owner(self, object, field) {
                let class_idx = self.str_const(&class_name);
                self.emit_u16(Op::GLOBAL_GET, class_idx);
                self.emit_u16(Op::STRUCT_GET, prop);
                let fn_tmp = self.define_local("__go_pending_instance_fn");
                self.emit_u16(Op::LOCAL_SET, fn_tmp);
                self.emit(Op::DROP);

                let receiver_slot = if self.pending_classes.get(&class_name).is_some_and(|pending| {
                    !pending
                        .instance_pointer_method_names
                        .iter()
                        .any(|name| self.canon(name) == self.canon(field))
                        && !pending.fields.is_empty()
                }) {
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit_autoderef_pointer_cell();
                    self.emit_user_value_type_clone_from_stack(&class_name);
                    let receiver_slot = self.define_local("__go_value_receiver");
                    self.emit_u16(Op::LOCAL_SET, receiver_slot);
                    self.emit(Op::DROP);
                    receiver_slot
                } else {
                    obj_tmp
                };

                let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                for (index, arg) in arg_exprs.iter().enumerate() {
                    self.compile_expr(arg)?;
                    let arg_slot = self.define_local(&format!("__go_pending_method_arg_{}", index));
                    self.emit_u16(Op::LOCAL_SET, arg_slot);
                    self.emit(Op::DROP);
                    arg_slots.push(arg_slot);
                }
                self.emit_call_ref_with_arg_slots(fn_tmp, Some(receiver_slot), &arg_slots);
                return Ok(());
            }

            self.emit_u16(Op::LOCAL_GET, obj_tmp);
            self.emit_u16(Op::STRUCT_GET, prop);
            let fn_tmp = self.define_local("__fn");
            self.emit_u16(Op::LOCAL_SET, fn_tmp); self.emit(Op::DROP);
            if self.profile.name == "php" {
                if let Some(overload) = self.resolve_instance_method_overload(object, field, &arg_exprs, false) {
                    let line = self.line;
                    self.emit_u16(Op::REF_FUNC, overload.chunk_idx as u16);
                    self.chunk().emit(0, line);
                    let direct_fn_tmp = self.define_local("__php_direct_instance_fn");
                    self.emit_u16(Op::LOCAL_SET, direct_fn_tmp);
                    self.emit(Op::DROP);
                    if overload.signature.has_rest {
                        self.emit_known_rest_call_from_local(direct_fn_tmp, Some(obj_tmp), args, &overload.signature)?;
                    } else {
                        self.emit_u16(Op::LOCAL_GET, direct_fn_tmp);
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        for a in &arg_exprs { self.compile_expr(a)?; }
                        self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                    }
                    return Ok(());
                }
            }
            let member_index_done = if self.profile.parens_for_index && !arg_exprs.is_empty() {
                self.emit_u16(Op::LOCAL_GET, fn_tmp);
                self.emit(Op::REF_TYPEOF);
                self.emit_const(Value::String(Arc::from("function")));
                self.emit(Op::DYN_EQ);
                let use_call = self.emit_jump(Op::BR_IF_TRUE);

                self.emit_u16(Op::LOCAL_GET, fn_tmp);
                for arg in &arg_exprs {
                    self.compile_expr(arg)?;
                    let line = self.line;
                    common::collections::emit_get(&mut self.chunks, self.current, line);
                }
                let done = self.emit_jump(Op::BR);
                self.patch_jump(use_call);
                Some(done)
            } else {
                None
            };
            if resolves_to_static_container_method(self, object, field) {
                if self.profile.name == "php" {
                    let class_canon = self.canon(&self.flatten_member_chain(object).join("."));
                    if let Some(overload) = self.resolve_static_method_overload_for_type(&class_canon, field, &arg_exprs) {
                        let line = self.line;
                        self.emit_u16(Op::REF_FUNC, overload.chunk_idx as u16);
                        self.chunk().emit(0, line);
                        let direct_fn_tmp = self.define_local("__php_direct_static_fn");
                        self.emit_u16(Op::LOCAL_SET, direct_fn_tmp);
                        self.emit(Op::DROP);
                        if overload.signature.has_rest {
                            self.emit_known_rest_call_from_local(direct_fn_tmp, Some(obj_tmp), args, &overload.signature)?;
                        } else {
                            self.emit_u16(Op::LOCAL_GET, direct_fn_tmp);
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            for a in &arg_exprs { self.compile_expr(a)?; }
                            self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                        }
                        if let Some(done) = member_index_done {
                            self.patch_jump(done);
                        }
                        return Ok(());
                    }
                }
                if self.profile.name == "csharp" && args.len() == 1 && !args[0].spread {
                    let class_canon = self.canon(&self.flatten_member_chain(object).join("."));
                    if self.resolve_static_method_overload_for_type(&class_canon, field, &arg_exprs)
                        .is_some_and(|overload| overload.signature.has_rest)
                    {
                        self.emit_variadic_array_call_from_local(fn_tmp, &args[0].value)?;
                        if let Some(done) = member_index_done {
                            self.patch_jump(done);
                        }
                        return Ok(());
                    }
                }
                let static_overload = self
                    .flatten_member_chain(object)
                    .last()
                    .and_then(|class_name| self.resolve_static_method_overload_for_type(class_name, field, &arg_exprs));
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
                        self.emit(Op::DROP);
                        direct_fn_tmp
                    } else {
                        fn_tmp
                    };
                    self.emit_known_rest_call_from_local(
                        rest_callee_slot,
                        if self.profile.name == "php" { Some(obj_tmp) } else { None },
                        args,
                        signature,
                    )?;
                } else {
                    let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                    for (index, arg) in arg_exprs.iter().enumerate() {
                        self.compile_expr(arg)?;
                        let arg_slot = self.define_local(&format!("__static_member_call_arg_{}", index));
                        self.emit_u16(Op::LOCAL_SET, arg_slot);
                        self.emit(Op::DROP);
                        arg_slots.push(arg_slot);
                    }
                    self.emit_call_ref_with_arg_slots(
                        fn_tmp,
                        if self.profile.name == "php" { Some(obj_tmp) } else { None },
                        &arg_slots,
                    );
                }
                if args.iter().any(|arg| arg.by_ref) {
                    let pack_slot = self.define_local("__member_static_container_pack");
                    self.emit_u16(Op::LOCAL_SET, pack_slot);
                    self.emit(Op::DROP);
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
                if let Some(done) = member_index_done {
                    self.patch_jump(done);
                }
                return Ok(());
            }
            let primitive_tostring_done = if self.profile.namespaces.use_dotnet
                && arg_exprs.is_empty()
                && field.eq_ignore_ascii_case("ToString")
            {
                let type_tmp = self.define_local("__dotnet_tostring_type");
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit(Op::REF_TYPEOF);
                self.emit_u16(Op::LOCAL_SET, type_tmp);
                self.emit(Op::DROP);

                let mut primitive_matches = Vec::new();
                for type_name in ["number", "i32", "i64", "string", "boolean"] {
                    self.emit_u16(Op::LOCAL_GET, type_tmp);
                    self.emit_const(Value::String(Arc::from(type_name)));
                    self.emit(Op::DYN_EQ);
                    primitive_matches.push(self.emit_jump(Op::BR_IF_TRUE));
                }
                let not_primitive = self.emit_jump(Op::BR);

                for patch in primitive_matches {
                    self.patch_jump(patch);
                }
                let tostring_global = self.str_const("__vybe_tostring");
                self.emit_u16(Op::GLOBAL_GET, tostring_global);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u8(Op::CALL_REF, 1);
                let done = self.emit_jump(Op::BR);

                self.patch_jump(not_primitive);
                Some(done)
            } else {
                if let Some(done) = member_index_done {
                    self.patch_jump(done);
                }
                None
            };
            if self.profile.namespaces.use_dotnet
                && arg_exprs.is_empty()
                && field.eq_ignore_ascii_case("Count")
            {
                self.emit_u16(Op::LOCAL_GET, fn_tmp);
                self.emit(Op::REF_TYPEOF);
                self.emit_const(Value::String(Arc::from("function")));
                self.emit(Op::DYN_EQ);
                let return_value = self.emit_jump(Op::BR_IF_FALSE);

                self.emit_u16(Op::LOCAL_GET, fn_tmp);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_u8(Op::CALL_REF, 1);
                let done = self.emit_jump(Op::BR);

                self.patch_jump(return_value);
                self.emit_u16(Op::LOCAL_GET, fn_tmp);
                self.patch_jump(done);
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
                self.emit(Op::DROP);
                self.emit_known_rest_call_from_local(
                    direct_fn_tmp,
                    if self.is_js_profile() { None } else { Some(obj_tmp) },
                    args,
                    &overload.signature,
                )?;
            } else {
                if self.is_js_profile() {
                    let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                    for (index, arg) in arg_exprs.iter().enumerate() {
                        self.compile_expr(arg)?;
                        let arg_slot = self.define_local(&format!("__js_member_call_arg_{}", index));
                        self.emit_u16(Op::LOCAL_SET, arg_slot);
                        self.emit(Op::DROP);
                        arg_slots.push(arg_slot);
                    }
                    self.emit_call_ref_with_arg_slots(fn_tmp, None, &arg_slots);
                } else {
                    let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                    for (index, arg) in arg_exprs.iter().enumerate() {
                        self.compile_expr(arg)?;
                        let arg_slot = self.define_local(&format!("__member_call_arg_{}", index));
                        self.emit_u16(Op::LOCAL_SET, arg_slot);
                        self.emit(Op::DROP);
                        arg_slots.push(arg_slot);
                    }
                    self.emit_call_ref_with_arg_slots(fn_tmp, Some(obj_tmp), &arg_slots);
                }
            }
            if let Some(done) = member_index_done {
                self.patch_jump(done);
            }
            if let Some(done) = primitive_tostring_done {
                self.patch_jump(done);
            }
            return Ok(());
        }

        // ── Simple call: name(args) / expr(args) ────────────────────
        if let ExprKind::Ident(name) = &callee.kind {
            let rest_signature = self
                .function_signatures
                .get(&self.canon(name))
                .and_then(|signatures| self.select_call_signature(signatures, args))
                .filter(|signature| signature.has_rest)
                .cloned();

            if self.is_php_profile()
                && (name.eq_ignore_ascii_case("exit") || name.eq_ignore_ascii_case("die"))
            {
                if let Some(arg) = arg_exprs.first() {
                    let arg_slot = self.define_local("__php_exit_arg");
                    self.compile_expr(arg)?;
                    self.emit_u16(Op::LOCAL_SET, arg_slot);
                    self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, arg_slot);
                    let typeof_idx = self.import("ecma:value", "typeof");
                    self.emit_host_call(typeof_idx, 1);
                    self.emit_const(Value::String(Arc::from("string")));
                    self.emit(Op::DYN_EQ);
                    let skip_print = self.emit_jump(Op::BR_IF_FALSE);

                    let log_idx = self.import("wasi:cli", "log");
                    let line = self.line;
                    self.emit_u16(Op::LOCAL_GET, arg_slot);
                    self.emit_common("php.echo_stringify", 1, line);
                    common::io::emit_print_with_import(self.chunk(), log_idx, 1, line);

                    self.patch_jump(skip_print);
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
                        self.emit_u16(Op::LOCAL_SET, fn_tmp); self.emit(Op::DROP);

                        let method_canon = self.canon(name);
                        if let Some(param_modes) = self.function_param_modes.get(&method_canon).cloned() {
                            if param_modes.iter().any(|mode| matches!(mode, PassBy::Ref | PassBy::Out)) {
                                let mut arg_slots = Vec::with_capacity(args.len());
                                for (index, arg) in args.iter().enumerate() {
                                    match param_modes.get(index).copied().unwrap_or(PassBy::Value) {
                                        PassBy::Out => self.emit(Op::NULL),
                                        PassBy::Ref | PassBy::Const | PassBy::Value => {
                                            self.compile_expr_with_value_copy(&arg.value)?;
                                        }
                                    }
                                    let arg_slot = self.define_local(&format!("__bare_static_call_arg_{}", index));
                                    self.emit_u16(Op::LOCAL_SET, arg_slot);
                                    self.emit(Op::DROP);
                                    arg_slots.push(arg_slot);
                                }

                                self.emit_u16(Op::LOCAL_GET, fn_tmp);
                                for slot in &arg_slots {
                                    self.emit_u16(Op::LOCAL_GET, *slot);
                                }
                                self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);

                                let pack_slot = self.define_local("__bare_static_ref_call_pack");
                                self.emit_u16(Op::LOCAL_SET, pack_slot);
                                self.emit(Op::DROP);
                                let mut ref_out_index = 1usize;
                                for (index, arg) in args.iter().enumerate() {
                                    if !matches!(param_modes.get(index), Some(PassBy::Ref | PassBy::Out)) {
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

                        if self.profile.name == "csharp" && args.len() == 1 && !args[0].spread {
                            if self.resolve_static_method_overload_for_type(&class_name, name, &arg_exprs)
                                .is_some_and(|overload| overload.signature.has_rest)
                            {
                                self.emit_variadic_array_call_from_local(fn_tmp, &args[0].value)?;
                                return Ok(());
                            }
                        }

                        let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                        for (index, arg) in arg_exprs.iter().enumerate() {
                            self.compile_expr_with_value_copy(arg)?;
                            let arg_slot = self.define_local(&format!("__bare_static_call_arg_{}", index));
                            self.emit_u16(Op::LOCAL_SET, arg_slot);
                            self.emit(Op::DROP);
                            arg_slots.push(arg_slot);
                        }
                        self.emit_call_ref_with_arg_slots(fn_tmp, None, &arg_slots);
                        return Ok(());
                    }
                }
            }

            let is_known_func = self.defined_functions.contains(name)
                || (!self.case_sensitive && self.defined_functions.iter().any(|g| g.eq_ignore_ascii_case(name)));
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
            if !is_known_func && arg_exprs.len() == 1 && self.profile.parens_for_index && !self.is_php_profile() {
                let canon_name = self.canon(name);
                let is_local = self.has_accessible_local_binding(name);
                let is_global_var = self.defined_globals.contains(&canon_name)
                    && !self.defined_classes.contains(&canon_name)
                    && !self.defined_functions.contains(&canon_name);
                let is_callable_typed = self
                    .lookup_var_type_hint(name)
                    .is_some_and(Self::is_callable_type_hint);
                if (is_local || is_global_var) && !is_callable_typed {
                    self.emit_var_get(name);
                    self.compile_expr(arg_exprs[0])?;
                    { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
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
                        self.emit(Op::DUP);
                        self.emit_u16(Op::STRUCT_GET, prop);
                        let fn_tmp = self.define_local("__bare_fn");
                        self.emit_u16(Op::LOCAL_SET, fn_tmp); self.emit(Op::DROP);
                        let obj_tmp = self.define_local("__bare_obj");
                        self.emit_u16(Op::LOCAL_SET, obj_tmp); self.emit(Op::DROP);

                        let is_callable_field = self.lookup_implicit_self_field_type_hint(name)
                            .is_some_and(Self::is_callable_type_hint);
                        if (self.profile.name == "pascal" && self.is_class_field(name)) || is_callable_field {
                            let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                            for (index, arg) in arg_exprs.iter().enumerate() {
                                self.compile_expr_with_value_copy(arg)?;
                                let arg_slot = self.define_local(&format!("__bare_field_call_arg_{}", index));
                                self.emit_u16(Op::LOCAL_SET, arg_slot);
                                self.emit(Op::DROP);
                                arg_slots.push(arg_slot);
                            }
                            self.emit_call_ref_with_arg_slots(fn_tmp, None, &arg_slots);
                            return Ok(());
                        }

                        let mut arg_slots = Vec::with_capacity(arg_exprs.len());
                        for (index, arg) in arg_exprs.iter().enumerate() {
                            self.compile_expr_with_value_copy(arg)?;
                            let arg_slot = self.define_local(&format!("__bare_method_call_arg_{}", index));
                            self.emit_u16(Op::LOCAL_SET, arg_slot);
                            self.emit(Op::DROP);
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
                    self.emit(Op::DROP);
                    self.emit_known_rest_call_from_local(callee_slot, None, args, signature)?;
                    return Ok(());
                }

                if self.profile.name == "php" && args.len() == 1 && args[0].spread {
                    if let Some(signature) = self
                        .function_signatures
                        .get(&self.canon(name))
                        .and_then(|signatures| signatures.iter().find(|signature| !signature.has_rest && !signature.param_names.is_empty()))
                        .cloned()
                    {
                        let spread_slot = self.define_local("__php_named_unpack");
                        self.compile_expr(&args[0].value)?;
                        self.emit_u16(Op::LOCAL_SET, spread_slot);
                        self.emit(Op::DROP);

                        let probe_slot = self.define_local("__php_named_unpack_probe");
                        self.emit_u16(Op::LOCAL_GET, spread_slot);
                        self.emit_const(Value::String(Arc::from(signature.param_names[0].as_str())));
                        common::collections::emit_get(&mut self.chunks, self.current, self.line);
                        self.emit_u16(Op::LOCAL_SET, probe_slot);
                        self.emit(Op::DROP);

                        self.emit_u16(Op::LOCAL_GET, probe_slot);
                        let positional_fallback = self.emit_jump(Op::BR_IF_NULL);

                        self.emit_var_get(name);
                        for param_name in &signature.param_names {
                            self.emit_u16(Op::LOCAL_GET, spread_slot);
                            self.emit_const(Value::String(Arc::from(param_name.as_str())));
                            common::collections::emit_get(&mut self.chunks, self.current, self.line);
                        }
                        self.emit_u8(Op::CALL_REF, signature.param_names.len() as u8);

                        let done = self.emit_jump(Op::BR);
                        self.patch_jump(positional_fallback);

                        let line = self.line;
                        let args_slot = self.define_local("__spread_args");
                        common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
                        self.emit_u16(Op::LOCAL_SET, args_slot); self.emit(Op::DROP);
                        let mut known_len: Option<usize> = Some(0);
                        for a in args {
                            if a.spread {
                                self.emit_u16(Op::LOCAL_GET, args_slot);
                                self.compile_expr(&a.value)?;
                                common::collections::emit_concat(&mut self.chunks, self.current, line);
                                self.emit_u16(Op::LOCAL_SET, args_slot); self.emit(Op::DROP);
                                if let ExprKind::Array(elems) = &a.value.kind {
                                    if let Some(ref mut k) = known_len { *k += elems.len(); }
                                } else {
                                    known_len = None;
                                }
                            } else {
                                self.emit_u16(Op::LOCAL_GET, args_slot);
                                self.compile_expr(&a.value)?;
                                common::collections::emit_push(&mut self.chunks, self.current, line);
                                self.emit(Op::DROP);
                                if let Some(ref mut k) = known_len { *k += 1; }
                            }
                        }
                        if let Some(arity) = known_len {
                            self.emit_var_get(name);
                            self.emit_u16(Op::LOCAL_GET, args_slot);
                            self.emit(Op::SPREAD);
                            self.emit_u8(Op::CALL_REF, arity as u8);
                        } else {
                            self.emit_u16(Op::LOCAL_GET, args_slot);
                            self.emit_const(Value::I32(16));
                            common::collections::emit_new_with_length(&mut self.chunks, self.current, line);
                            common::collections::emit_concat(&mut self.chunks, self.current, line);
                            self.emit_const(Value::F64(0.0));
                            self.emit_const(Value::F64(16.0));
                            common::collections::emit_slice(&mut self.chunks, self.current, line);
                            self.emit_u16(Op::LOCAL_SET, args_slot); self.emit(Op::DROP);

                            self.emit_var_get(name);
                            self.emit_u16(Op::LOCAL_GET, args_slot);
                            self.emit(Op::SPREAD);
                            self.emit_u8(Op::CALL_REF, 16);
                        }

                        self.patch_jump(done);
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
                self.emit_u16(Op::LOCAL_SET, args_slot); self.emit(Op::DROP);
                let mut known_len: Option<usize> = Some(0);
                for a in args {
                    if a.spread {
                        // new_arr = concat(args, spread)
                        self.emit_u16(Op::LOCAL_GET, args_slot);
                        self.compile_expr(&a.value)?;
                        common::collections::emit_concat(&mut self.chunks, self.current, line);
                        self.emit_u16(Op::LOCAL_SET, args_slot); self.emit(Op::DROP);
                        if let ExprKind::Array(elems) = &a.value.kind {
                            if let Some(ref mut k) = known_len { *k += elems.len(); }
                        } else {
                            known_len = None;
                        }
                    } else {
                        self.emit_u16(Op::LOCAL_GET, args_slot);
                        self.compile_expr(&a.value)?;
                        common::collections::emit_push(&mut self.chunks, self.current, line);
                        self.emit(Op::DROP); // drop new_length returned by push
                        if let Some(ref mut k) = known_len { *k += 1; }
                    }
                }
                self.emit_var_get(name);
                let callee_slot = self.define_local("__ident_spread_callee");
                self.emit_u16(Op::LOCAL_SET, callee_slot);
                self.emit(Op::DROP);
                self.emit_php_dynamic_function_name_resolution(callee_slot);
                let receiver_key = self.str_const("__vybe_method_receiver");
                self.emit_u16(Op::LOCAL_GET, callee_slot);
                self.emit_u16(Op::STRUCT_GET, receiver_key);
                let receiver_slot = self.define_local("__ident_spread_receiver");
                self.emit_u16(Op::LOCAL_SET, receiver_slot);
                self.emit(Op::DROP);
                self.emit_call_ref_with_args_array(callee_slot, Some(receiver_slot), args_slot, known_len);
                return Ok(());
            }
            if self.is_python_profile() && !is_known_func {
                let callee_slot = self.define_local("__py_call_target");
                self.emit_var_get(name);
                self.emit_u16(Op::LOCAL_SET, callee_slot);
                self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, callee_slot);
                let typeof_idx = self.import("ecma:value", "typeof");
                self.emit_host_call(typeof_idx, 1);
                self.emit_const(Value::String(Arc::from("function")));
                self.emit(Op::DYN_EQ);
                let invoke_dunder = self.emit_jump(Op::BR_IF_FALSE);

                self.emit_u16(Op::LOCAL_GET, callee_slot);
                for a in &arg_exprs { self.compile_expr(a)?; }
                self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                let end = self.emit_jump(Op::BR);

                self.patch_jump(invoke_dunder);
                self.emit_u16(Op::LOCAL_GET, callee_slot);
                let call_prop = self.str_const("call");
                self.emit_u16(Op::STRUCT_GET, call_prop);
                let call_slot = self.define_local("__py_call_method");
                self.emit_u16(Op::LOCAL_SET, call_slot);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, call_slot);
                self.emit(Op::REF_IS_NULL);
                let try_dunder_name = self.emit_jump(Op::BR_IF_TRUE);
                self.emit_u16(Op::LOCAL_GET, call_slot);
                self.emit_u16(Op::LOCAL_GET, callee_slot);
                for a in &arg_exprs { self.compile_expr(a)?; }
                self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                let found_end = self.emit_jump(Op::BR);

                self.patch_jump(try_dunder_name);
                self.emit_u16(Op::LOCAL_GET, callee_slot);
                let dunder_prop = self.str_const("__call__");
                self.emit_u16(Op::STRUCT_GET, dunder_prop);
                let dunder_slot = self.define_local("__py_dunder_call_method");
                self.emit_u16(Op::LOCAL_SET, dunder_slot);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, dunder_slot);
                self.emit(Op::REF_IS_NULL);
                let no_dunder = self.emit_jump(Op::BR_IF_TRUE);
                self.emit_u16(Op::LOCAL_GET, dunder_slot);
                self.emit_u16(Op::LOCAL_GET, callee_slot);
                for a in &arg_exprs { self.compile_expr(a)?; }
                self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                let dunder_end = self.emit_jump(Op::BR);

                self.patch_jump(no_dunder);
                self.emit(Op::UNDEFINED);
                self.patch_jump(found_end);
                self.patch_jump(dunder_end);
                self.patch_jump(end);
                return Ok(());
            }

            let callee_slot = self.define_local("__direct_call_callee");
            self.emit_var_get(name);
            self.emit_u16(Op::LOCAL_SET, callee_slot);
            self.emit(Op::DROP);
            self.emit_php_dynamic_function_name_resolution(callee_slot);
            let receiver_key = self.str_const("__vybe_method_receiver");
            self.emit_u16(Op::LOCAL_GET, callee_slot);
            self.emit_u16(Op::STRUCT_GET, receiver_key);
            let receiver_slot = self.define_local("__direct_call_receiver");
            self.emit_u16(Op::LOCAL_SET, receiver_slot);
            self.emit(Op::DROP);
            if let Some(signature) = rest_signature.as_ref() {
                self.emit_u16(Op::LOCAL_GET, receiver_slot);
                self.emit(Op::REF_IS_NULL);
                let no_receiver = self.emit_jump(Op::BR_IF_TRUE);

                self.emit_known_rest_call_from_local(callee_slot, Some(receiver_slot), args, signature)?;
                let call_done = self.emit_jump(Op::BR);

                self.patch_jump(no_receiver);
                self.emit_known_rest_call_from_local(callee_slot, None, args, signature)?;
                self.patch_jump(call_done);
                return Ok(());
            }
            if let Some(param_modes) = self.function_param_modes.get(&self.canon(name)).cloned() {
                if param_modes.iter().any(|mode| matches!(mode, PassBy::Ref | PassBy::Out)) {
                    let mut arg_slots = Vec::with_capacity(args.len());
                    for (index, arg) in args.iter().enumerate() {
                        match param_modes.get(index).copied().unwrap_or(PassBy::Value) {
                            PassBy::Out => self.emit(Op::NULL),
                            PassBy::Ref | PassBy::Const | PassBy::Value => {
                                if !matches!(param_modes.get(index), Some(PassBy::Out)) {
                                    self.compile_expr_with_value_copy(&arg.value)?;
                                }
                            }
                        }

                        let arg_slot = self.define_local(&format!("__direct_call_arg_{}", index));
                        self.emit_u16(Op::LOCAL_SET, arg_slot);
                        self.emit(Op::DROP);
                        arg_slots.push(arg_slot);
                    }

                    self.emit_u16(Op::LOCAL_GET, callee_slot);
                    self.emit_u16(Op::LOCAL_GET, receiver_slot);
                    self.emit(Op::REF_IS_NULL);
                    let no_receiver = self.emit_jump(Op::BR_IF_TRUE);

                    self.emit_u16(Op::LOCAL_GET, callee_slot);
                    self.emit_u16(Op::LOCAL_GET, receiver_slot);
                    for slot in &arg_slots {
                        self.emit_u16(Op::LOCAL_GET, *slot);
                    }
                    self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                    let call_done = self.emit_jump(Op::BR);

                    self.patch_jump(no_receiver);
                    self.emit_u16(Op::LOCAL_GET, callee_slot);
                    for slot in &arg_slots {
                        self.emit_u16(Op::LOCAL_GET, *slot);
                    }
                    self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                    self.patch_jump(call_done);

                    let pack_slot = self.define_local("__ref_call_pack");
                    self.emit_u16(Op::LOCAL_SET, pack_slot);
                    self.emit(Op::DROP);
                    let mut ref_out_index = 1usize;
                    for (index, arg) in args.iter().enumerate() {
                        if !matches!(param_modes.get(index), Some(PassBy::Ref | PassBy::Out)) {
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
                self.compile_expr_with_value_copy(arg)?;
                let arg_slot = self.define_local(&format!("__direct_call_arg_{}", index));
                self.emit_u16(Op::LOCAL_SET, arg_slot);
                self.emit(Op::DROP);
                arg_slots.push(arg_slot);
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
                let obj_tmp = self.define_local("__js_idx_obj");
                self.compile_expr(object)?;
                self.emit_u16(Op::LOCAL_SET, obj_tmp); self.emit(Op::DROP);
                let saved_js_this = self.save_js_this("__js_prev_this_idx");
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.set_js_this_from_stack();
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.compile_expr(index)?;
                let line = self.line;
                common::collections::emit_get(&mut self.chunks, self.current, line);
                for a in &arg_exprs { self.compile_expr(a)?; }
                self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                let result_slot = self.define_local("__js_idx_result");
                self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                self.restore_js_this(saved_js_this);
                self.emit_u16(Op::LOCAL_GET, result_slot);
                return Ok(());
            }
        }

        if self.profile.parens_for_index && !arg_exprs.is_empty() {
            let is_bound_array = matches!(&callee.kind,
                ExprKind::Ident(name) if self.lookup_array_binding(name).is_some()
            );
            let is_indexable_typed = is_bound_array || self
                .infer_expr_type_hint(callee)
                .as_deref()
                .map(Self::normalize_type_hint)
                .is_some_and(|type_hint| type_hint.ends_with("()") && !Self::is_callable_type_hint(&type_hint));
            if is_indexable_typed {
                self.compile_expr(callee)?;
                for arg in &arg_exprs {
                    self.compile_expr(arg)?;
                    let line = self.line;
                    common::collections::emit_get(&mut self.chunks, self.current, line);
                }
                return Ok(());
            }
        }

        let mut runtime_index_done: Option<usize> = None;

        // ── Fallback: general expression call ───────────────────────
        self.compile_expr(callee)?;
        let callee_slot = self.define_local("__call_ref_callee");
        self.emit_u16(Op::LOCAL_SET, callee_slot);
        self.emit(Op::DROP);
        self.emit_php_dynamic_function_name_resolution(callee_slot);

        if self.profile.parens_for_index
            && !arg_exprs.is_empty()
            && matches!(&callee.kind, ExprKind::Call { .. } | ExprKind::Index { .. })
        {
            self.emit_u16(Op::LOCAL_GET, callee_slot);
            self.emit(Op::REF_IS_ARRAY);
            let not_runtime_array = self.emit_jump(Op::BR_IF_FALSE);

            self.emit_u16(Op::LOCAL_GET, callee_slot);
            for arg in &arg_exprs {
                self.compile_expr(arg)?;
                let line = self.line;
                common::collections::emit_get(&mut self.chunks, self.current, line);
            }
            runtime_index_done = Some(self.emit_jump(Op::BR));
            self.patch_jump(not_runtime_array);
        }

        let has_by_ref_args = args.iter().any(|arg| arg.by_ref);
        let receiver_key = self.str_const("__vybe_method_receiver");
        self.emit_u16(Op::LOCAL_GET, callee_slot);
        self.emit_u16(Op::STRUCT_GET, receiver_key);
        let receiver_slot = self.define_local("__call_ref_receiver");
        self.emit_u16(Op::LOCAL_SET, receiver_slot);
        self.emit(Op::DROP);

        let mut arg_slots = Vec::with_capacity(arg_exprs.len());
        for (index, arg) in arg_exprs.iter().enumerate() {
            self.compile_expr(arg)?;
            let arg_slot = self.define_local(&format!("__call_ref_arg_{}", index));
            self.emit_u16(Op::LOCAL_SET, arg_slot);
            self.emit(Op::DROP);
            arg_slots.push(arg_slot);
        }

        self.emit_u16(Op::LOCAL_GET, receiver_slot);
        self.emit(Op::REF_IS_NULL);
        let no_receiver = self.emit_jump(Op::BR_IF_TRUE);

        self.emit_u16(Op::LOCAL_GET, callee_slot);
        self.emit_u16(Op::LOCAL_GET, receiver_slot);
        for slot in &arg_slots {
            self.emit_u16(Op::LOCAL_GET, *slot);
        }
        self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
        let result_slot = self.define_local("__call_ref_result");
        self.emit_u16(Op::LOCAL_SET, result_slot);
        self.emit(Op::DROP);
        let done = self.emit_jump(Op::BR);

        self.patch_jump(no_receiver);
        self.emit_u16(Op::LOCAL_GET, callee_slot);
        for slot in &arg_slots {
            self.emit_u16(Op::LOCAL_GET, *slot);
        }
        self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
        self.emit_u16(Op::LOCAL_SET, result_slot);
        self.emit(Op::DROP);
        self.patch_jump(done);

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

        if let Some(index_done) = runtime_index_done {
            self.patch_jump(index_done);
        }
        Ok(())
    }

    fn try_compile_dotnet_guid_try_parse(&mut self, callee: &Expression, args: &[Argument]) -> Result<bool, String> {
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
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, parsed_slot);
        self.emit(Op::REF_IS_NULL);
        let invalid = self.emit_jump(Op::BR_IF_TRUE);

        self.emit_u16(Op::LOCAL_GET, parsed_slot);
        self.compile_assign_target(&args[1].value)?;
        if let ExprKind::Ident(name) = &args[1].value.kind {
            let normalized = Self::normalize_type_hint("Guid");
            if let Some(slot) = self.scope().resolve_ci(name) {
                if let Some(local) = self.scope_mut().locals.iter_mut().rev().find(|local| local.slot == slot) {
                    local.type_hint = Some(normalized.clone());
                }
            } else {
                self.global_type_hints.insert(self.canon(name), normalized.clone());
            }
        }
        self.emit(Op::TRUE);
        let done = self.emit_jump(Op::BR);

        self.patch_jump(invalid);
        self.emit(Op::NULL);
        self.compile_assign_target(&args[1].value)?;
        self.emit(Op::FALSE);
        self.patch_jump(done);
        Ok(true)
    }

    fn try_compile_dotnet_numeric_try_parse(&mut self, callee: &Expression, args: &[Argument]) -> Result<bool, String> {
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

        let line = self.line;
        self.compile_expr(&args[0].value)?;
        let number_idx = self.import("ecma:number", "Number");
        self.emit_host_call(number_idx, 1);

        let parsed_slot = self.define_local("__numeric_try_parse_value");
        self.emit_u16(Op::LOCAL_SET, parsed_slot);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, parsed_slot);
        self.emit_u16(Op::LOCAL_GET, parsed_slot);
        self.emit(Op::DYN_EQ);
        let invalid = self.emit_jump(Op::BR_IF_FALSE);

        self.emit_u16(Op::LOCAL_GET, parsed_slot);
        self.emit(Op::F64_FLOOR);
        self.compile_assign_target(&args[1].value)?;
        if let ExprKind::Ident(name) = &args[1].value.kind {
            let normalized = Self::normalize_type_hint("int");
            if let Some(slot) = self.scope().resolve_ci(name) {
                if let Some(local) = self.scope_mut().locals.iter_mut().rev().find(|local| local.slot == slot) {
                    local.type_hint = Some(normalized.clone());
                }
            } else {
                self.global_type_hints.insert(self.canon(name), normalized);
            }
        }
        self.emit(Op::TRUE);
        let done = self.emit_jump(Op::BR);

        self.patch_jump(invalid);
        self.emit_const(Value::F64(0.0));
        self.compile_assign_target(&args[1].value)?;
        self.emit(Op::FALSE);
        self.patch_jump(done);
        let _ = line;
        Ok(true)
    }

    fn try_compile_dotnet_dictionary_try_get_value(&mut self, callee: &Expression, args: &[Argument]) -> Result<bool, String> {
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
                && (!self.case_sensitive && self.scope().resolve_ci(name).is_none() || self.case_sensitive)
                && !self.defined_globals.contains(&self.canon(name));
            if unresolved {
                self.define_local_typed(name, None);
            }
        }

        self.compile_expr(object)?;
        let map_slot = self.define_local("__dict_try_get_map");
        self.emit_u16(Op::LOCAL_SET, map_slot);
        self.emit(Op::DROP);

        self.compile_expr(&args[0].value)?;
        if self.expr_uses_case_insensitive_string_keys(object) {
            let line = self.line;
            common::strings::emit_to_lower(self.chunk(), line);
        }
        let key_slot = self.define_local("__dict_try_get_key");
        self.emit_u16(Op::LOCAL_SET, key_slot);
        self.emit(Op::DROP);

        let has_idx = self.import("ecma:map", "has");
        self.emit_u16(Op::LOCAL_GET, map_slot);
        self.emit_u16(Op::LOCAL_GET, key_slot);
        self.emit_host_call(has_idx, 2);
        let has_slot = self.define_local("__dict_try_get_has");
        self.emit_u16(Op::LOCAL_SET, has_slot);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, has_slot);
        let missing = self.emit_jump(Op::BR_IF_FALSE);

        self.emit_u16(Op::LOCAL_GET, map_slot);
        let getter_key = self.str_const("__get___index__");
        self.emit_u16(Op::STRUCT_GET, getter_key);
        let getter_slot = self.define_local("__dict_try_get_getter");
        self.emit_u16(Op::LOCAL_SET, getter_slot);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, getter_slot);
        self.emit(Op::REF_IS_NULL);
        let fallback = self.emit_jump(Op::BR_IF_TRUE);

        self.emit_u16(Op::LOCAL_GET, getter_slot);
        self.emit_u16(Op::LOCAL_GET, map_slot);
        self.emit_u16(Op::LOCAL_GET, key_slot);
        self.emit_u8(Op::CALL_REF, 2);
        let done = self.emit_jump(Op::BR);

        self.patch_jump(fallback);
        self.emit_u16(Op::LOCAL_GET, map_slot);
        self.emit_u16(Op::LOCAL_GET, key_slot);
        common::collections::emit_get(&mut self.chunks, self.current, self.line);
        self.patch_jump(done);

        self.compile_assign_target(&args[1].value)?;
        self.emit(Op::TRUE);
        let done = self.emit_jump(Op::BR);

        self.patch_jump(missing);
        self.emit(Op::NULL);
        self.compile_assign_target(&args[1].value)?;
        self.emit(Op::FALSE);
        self.patch_jump(done);
        Ok(true)
    }

    fn try_compile_dotnet_case_insensitive_collection_call(&mut self, callee: &Expression, args: &[Argument]) -> Result<bool, String> {
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
                    self.emit(Op::DROP);

                    self.compile_collection_key(object, &args[0].value)?;
                    self.emit_u16(Op::LOCAL_SET, key_slot);
                    self.emit(Op::DROP);

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
                    self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, keys_slot);
                    self.emit(Op::REF_IS_NULL);
                    let have_keys = self.emit_jump(Op::BR_IF_FALSE);

                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
                    self.emit(Op::DUP);
                    self.emit_u16(Op::LOCAL_SET, keys_slot);
                    self.emit(Op::DROP);
                    self.emit_u16(Op::STRUCT_SET, keys_key);
                    self.emit(Op::DROP);

                    self.patch_jump(have_keys);
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

    pub(crate) fn resolve_reflection_binding_expr(&self, expr: &Expression) -> Option<ReflectionBinding> {
        match &expr.kind {
            ExprKind::Lit(Literal::Str(type_name)) if type_name.starts_with("System.") => {
                Some(ReflectionBinding::Type(type_name.clone()))
            }
            ExprKind::Ident(name) => self.reflection_bindings.get(&self.canon(name)).cloned(),
            ExprKind::Member { object, field, .. } => {
                let receiver = self.resolve_reflection_binding_expr(object)?;
                match (receiver, strip_generic_suffix(field.as_str())) {
                    (ReflectionBinding::Type(type_name), "BaseType") => {
                        self.reflection_base_type_name(&type_name)
                            .map(ReflectionBinding::Type)
                    }
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
                        Some(ReflectionBinding::Method { type_name, method_name })
                    }
                    (ReflectionBinding::Type(type_name), "GetProperty") => {
                        let property_name = self.resolve_reflection_string_arg(args.first()?)?;
                        Some(ReflectionBinding::Property { type_name, property_name })
                    }
                    (ReflectionBinding::Type(type_name), "GetField") => {
                        let field_name = self.resolve_reflection_string_arg(args.first()?)?;
                        Some(ReflectionBinding::Field { type_name, field_name })
                    }
                    (ReflectionBinding::Type(type_name), "GetNestedType") => {
                        let nested_name = self.resolve_reflection_string_arg(args.first()?)?;
                        self.reflection_nested_type_name(&type_name, &nested_name)
                            .map(ReflectionBinding::Type)
                    }
                    (ReflectionBinding::Type(type_name), "GetGenericTypeDefinition") => {
                        Some(ReflectionBinding::Type(self.reflection_open_generic_type_name(&type_name)))
                    }
                    (ReflectionBinding::Type(type_name), "GetConstructor") => {
                        let param_types = self.resolve_reflection_type_array_expr(&args.first()?.value)?;
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
                let ExprKind::Member { object: method_object, field, .. } = &callee.kind else {
                    return None;
                };
                if strip_generic_suffix(field.as_str()) != "GetParameters" {
                    return None;
                }
                let ReflectionBinding::Method { type_name, method_name } = self.resolve_reflection_binding_expr(method_object)? else {
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

    fn resolve_reflection_string_arg(&self, arg: &Argument) -> Option<String> {
        match &arg.value.kind {
            ExprKind::Lit(Literal::Str(value)) => Some(value.clone()),
            ExprKind::Ident(name) => self.reflection_bindings.get(&self.canon(name)).and_then(|binding| {
                if let ReflectionBinding::Type(type_name) = binding {
                    Some(type_name.clone())
                } else {
                    None
                }
            }),
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
                items.iter()
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
        match (self.resolve_reflection_binding_expr(object)?, strip_generic_suffix(field.as_str())) {
            (ReflectionBinding::Type(type_name), "Name") => Some(self.reflection_type_short_name(&type_name)),
            (ReflectionBinding::Type(type_name), "FullName") => Some(self.reflection_type_full_name(&type_name)),
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
            self.emit(Op::DUP);
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
            ReflectionBinding::Type(type_name) => self.reflection_attributes_for_type(type_name, attribute_type, inherit),
            ReflectionBinding::Constructor { .. } => Vec::new(),
            ReflectionBinding::Method { type_name, method_name } => self
                .reflection_types
                .get(type_name)
                .and_then(|meta| meta.methods.get(method_name))
                .map(|meta| self.filter_reflection_attributes(&meta.decorators, attribute_type))
                .unwrap_or_default(),
            ReflectionBinding::Property { type_name, property_name } => self
                .reflection_types
                .get(type_name)
                .and_then(|meta| meta.properties.get(property_name))
                .map(|meta| self.filter_reflection_attributes(&meta.decorators, attribute_type))
                .unwrap_or_default(),
            ReflectionBinding::Field { type_name, field_name } => self
                .reflection_types
                .get(type_name)
                .and_then(|meta| meta.fields.get(field_name))
                .map(|meta| self.filter_reflection_attributes(&meta.decorators, attribute_type))
                .unwrap_or_default(),
            ReflectionBinding::Parameter { type_name, method_name, index } => self
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
                    let usage = self.attribute_usage.get(attribute_type).copied().unwrap_or_default();
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

        let positional_args: Vec<Argument> = args.iter().filter(|arg| arg.name.is_none()).cloned().collect();
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
        self.emit(Op::DROP);

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
            self.emit(Op::DUP);
            self.compile_reflection_attribute_instance(attr)?;
            common::collections::emit_push(&mut self.chunks, self.current, line);
            self.emit(Op::DROP);
        }
        Ok(())
    }

    fn compile_reflection_binding_value(&mut self, binding: &ReflectionBinding) -> Result<(), String> {
        match binding {
            ReflectionBinding::Type(type_name) => {
                self.compile_reflection_type_value(type_name)?;
            }
            ReflectionBinding::Constructor { .. } => {
                self.compile_expr(&Expression::new(ExprKind::Object(vec![
                    ObjectProperty::KeyValue {
                        key: Expression::string("Name"),
                        value: Expression::string(".ctor"),
                    },
                ])))?;
            }
            ReflectionBinding::Method { type_name, method_name } => {
                let is_static = self
                    .reflection_type_metadata(type_name)
                    .and_then(|meta| meta.methods.get(method_name))
                    .map(|meta| meta.is_static)
                    .unwrap_or(false);
                self.compile_expr(&Expression::new(ExprKind::Object(vec![
                    ObjectProperty::KeyValue {
                        key: Expression::string("Name"),
                        value: Expression::string(method_name),
                    },
                    ObjectProperty::KeyValue {
                        key: Expression::string("IsStatic"),
                        value: Expression::bool(is_static),
                    },
                ])))?;
            }
            ReflectionBinding::Property { type_name, property_name } => {
                let can_write = self
                    .reflection_type_metadata(type_name)
                    .and_then(|meta| meta.properties.get(property_name))
                    .map(|meta| meta.can_write)
                    .unwrap_or(false);
                self.compile_expr(&Expression::new(ExprKind::Object(vec![
                    ObjectProperty::KeyValue {
                        key: Expression::string("Name"),
                        value: Expression::string(property_name),
                    },
                    ObjectProperty::KeyValue {
                        key: Expression::string("CanWrite"),
                        value: Expression::bool(can_write),
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
            ReflectionBinding::Parameter { type_name, method_name, index } => {
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

    fn try_compile_dotnet_attribute_reflection_call(&mut self, callee: &Expression, args: &[Argument]) -> Result<bool, String> {
        let ExprKind::Member { object, field, .. } = &callee.kind else {
            return Ok(false);
        };
        let field_name = strip_generic_suffix(field);
        let receiver_type = terminal_type_name(object).unwrap_or_default();

        if (receiver_type.eq_ignore_ascii_case("Activator") || receiver_type.eq_ignore_ascii_case("System.Activator"))
            && field_name == "CreateInstance" && !args.is_empty()
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

        if (receiver_type.eq_ignore_ascii_case("Attribute") || receiver_type.eq_ignore_ascii_case("System.Attribute"))
            && field_name == "GetCustomAttribute" && args.len() >= 2
        {
            let Some(provider) = self.resolve_reflection_binding_expr(&args[0].value) else {
                return Ok(false);
            };
            let Some(attribute_type) = self.resolve_reflection_type_arg(&args[1].value) else {
                return Ok(false);
            };
            let attrs = self.reflection_attributes_for_binding(&provider, Some(&attribute_type), true);
            if let Some(attr) = attrs.first() {
                self.compile_reflection_attribute_instance(attr)?;
            } else {
                self.emit(Op::NULL);
            }
            return Ok(true);
        }

        if (receiver_type.eq_ignore_ascii_case("Attribute") || receiver_type.eq_ignore_ascii_case("System.Attribute"))
            && field_name == "IsDefined" && args.len() >= 2
        {
            let Some(provider) = self.resolve_reflection_binding_expr(&args[0].value) else {
                return Ok(false);
            };
            let Some(attribute_type) = self.resolve_reflection_type_arg(&args[1].value) else {
                return Ok(false);
            };
            let attrs = self.reflection_attributes_for_binding(&provider, Some(&attribute_type), true);
            self.emit(if attrs.is_empty() { Op::FALSE } else { Op::TRUE });
            return Ok(true);
        }

        let Some(provider) = self.resolve_reflection_binding_expr(object) else {
            return Ok(false);
        };
        match field_name {
            "GetMethod" if args.len() >= 1 => {
                let Some(binding) = self.resolve_reflection_binding_expr(&Expression::new(ExprKind::Call {
                    callee: Box::new(callee.clone()),
                    args: args.to_vec(),
                    optional: false,
                })) else {
                    return Ok(false);
                };
                self.compile_reflection_binding_value(&binding)?;
                Ok(true)
            }
            "GetProperty" if args.len() >= 1 => {
                let Some(binding) = self.resolve_reflection_binding_expr(&Expression::new(ExprKind::Call {
                    callee: Box::new(callee.clone()),
                    args: args.to_vec(),
                    optional: false,
                })) else {
                    return Ok(false);
                };
                self.compile_reflection_binding_value(&binding)?;
                Ok(true)
            }
            "GetField" if args.len() >= 1 => {
                let Some(binding) = self.resolve_reflection_binding_expr(&Expression::new(ExprKind::Call {
                    callee: Box::new(callee.clone()),
                    args: args.to_vec(),
                    optional: false,
                })) else {
                    return Ok(false);
                };
                self.compile_reflection_binding_value(&binding)?;
                Ok(true)
            }
            "GetConstructor" if args.len() >= 1 => {
                let Some(binding) = self.resolve_reflection_binding_expr(&Expression::new(ExprKind::Call {
                    callee: Box::new(callee.clone()),
                    args: args.to_vec(),
                    optional: false,
                })) else {
                    return Ok(false);
                };
                self.compile_reflection_binding_value(&binding)?;
                Ok(true)
            }
            "GetNestedType" if args.len() >= 1 => {
                let Some(binding) = self.resolve_reflection_binding_expr(&Expression::new(ExprKind::Call {
                    callee: Box::new(callee.clone()),
                    args: args.to_vec(),
                    optional: false,
                })) else {
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
                let Some(binding) = self.resolve_reflection_binding_expr(&Expression::new(ExprKind::Call {
                    callee: Box::new(callee.clone()),
                    args: args.to_vec(),
                    optional: false,
                })) else {
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
                self.emit(if self.reflection_is_assignable_from(&type_name, &other_type) {
                    Op::TRUE
                } else {
                    Op::FALSE
                });
                Ok(true)
            }
            "GetParameters" if args.is_empty() => {
                let ReflectionBinding::Method { type_name, method_name } = provider else {
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
                    self.emit(Op::DUP);
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
                let attrs = self.reflection_attributes_for_binding(&provider, Some(&attribute_type), inherit);
                self.compile_reflection_attribute_array(&attrs)?;
                Ok(true)
            }
            "Invoke" => {
                match provider {
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
                }
            }
            "GetValue" if !args.is_empty() => {
                match provider {
                    ReflectionBinding::Property { property_name, .. } | ReflectionBinding::Field { field_name: property_name, .. } => {
                        self.compile_expr(&Expression::new(ExprKind::Member {
                            object: Box::new(args[0].value.clone()),
                            field: property_name,
                            null_safe: false,
                        }))?;
                        Ok(true)
                    }
                    _ => Ok(false),
                }
            }
            "SetValue" if args.len() >= 2 => {
                match provider {
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
                            self.emit(Op::DROP);

                            self.compile_expr(&args[0].value)?;
                            self.emit(Op::DUP);
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
                }
            }
            _ => Ok(false),
        }
    }

    fn try_compile_dotnet_delegate_call(&mut self, callee: &Expression, args: &[Argument]) -> Result<bool, String> {
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

    fn try_compile_dotnet_formatted_tostring(&mut self, callee: &Expression, args: &[Argument]) -> Result<bool, String> {
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

    fn try_compile_dotnet_zero_arg_tostring(&mut self, callee: &Expression, args: &[Argument]) -> Result<bool, String> {
        if !matches!(self.profile.name.as_str(), "csharp" | "vb") || !self.profile.namespaces.use_dotnet {
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
            if let Some(target) = common::dotnet::surface().lookup_instance_method(&class_name, field, 0) {
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
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit(Op::REF_TYPEOF);
        let type_slot = self.define_local("__dotnet_tostring_type");
        self.emit_u16(Op::LOCAL_SET, type_slot);
        self.emit(Op::DROP);

        let mut primitive_matches = Vec::new();
        for type_name in ["number", "i32", "i64", "string", "boolean"] {
            self.emit_u16(Op::LOCAL_GET, type_slot);
            self.emit_const(Value::String(Arc::from(type_name)));
            self.emit(Op::DYN_EQ);
            primitive_matches.push(self.emit_jump(Op::BR_IF_TRUE));
        }
        let object_path = self.emit_jump(Op::BR);

        for patch in primitive_matches {
            self.patch_jump(patch);
        }
        let tostring_global = self.str_const("__vybe_tostring");
        self.emit_u16(Op::GLOBAL_GET, tostring_global);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_u8(Op::CALL_REF, 1);
        let done = self.emit_jump(Op::BR);

        self.patch_jump(object_path);
        let canon_key = self.str_const(&self.canon(field));
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_u16(Op::STRUCT_GET, canon_key);
        let fn_slot = self.define_local("__dotnet_tostring_fn");
        self.emit_u16(Op::LOCAL_SET, fn_slot);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, fn_slot);
        self.emit(Op::REF_IS_UNDEFINED);
        let have_canon = self.emit_jump(Op::BR_IF_FALSE);
        if field.as_str() != self.canon(field) {
            let exact_key = self.str_const(field);
            self.emit_u16(Op::LOCAL_GET, obj_slot);
            self.emit_u16(Op::STRUCT_GET, exact_key);
            self.emit_u16(Op::LOCAL_SET, fn_slot);
            self.emit(Op::DROP);
        }
        self.patch_jump(have_canon);

        self.emit_u16(Op::LOCAL_GET, fn_slot);
        self.emit(Op::REF_IS_UNDEFINED);
        let no_method = self.emit_jump(Op::BR_IF_FALSE);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        let type_key = self.str_const("__type");
        self.emit_u16(Op::STRUCT_GET, type_key);
        self.emit_const(Value::String(Arc::from("Guid")));
        self.emit(Op::DYN_EQ);
        let not_guid = self.emit_jump(Op::BR_IF_FALSE);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        let value_key = self.str_const("__value");
        self.emit_u16(Op::STRUCT_GET, value_key);
        let object_done = self.emit_jump(Op::BR);
        self.patch_jump(not_guid);
        self.emit_u16(Op::GLOBAL_GET, tostring_global);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_u8(Op::CALL_REF, 1);
        let fallback_done = self.emit_jump(Op::BR);

        self.patch_jump(no_method);
        self.emit_u16(Op::LOCAL_GET, fn_slot);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_u8(Op::CALL_REF, 1);
        self.patch_jump(object_done);
        self.patch_jump(fallback_done);
        self.patch_jump(done);
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
            ExprKind::Ident(name) => {
                self.lookup_var_type_hint(name)
                    .and_then(|hint| self.resolve_known_enum_type(hint))
                    .or_else(|| self.resolve_known_enum_type(name))
            }
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

    fn emit_enum_name_lookup(&mut self, enum_type: &str, value_expr: &Expression, ignore_case: bool) -> Result<(), String> {
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
        self.emit(Op::DROP);

        let mut done_jumps = Vec::new();
        for (_, name) in entries {
            self.emit_u16(Op::LOCAL_GET, input_slot);
            let candidate = if ignore_case { name.to_ascii_lowercase() } else { name.clone() };
            self.emit_const(Value::String(Arc::from(candidate.as_str())));
            self.emit(Op::DYN_EQ);
            let no_match = self.emit_jump(Op::BR_IF_FALSE);
            self.emit_const(Value::String(Arc::from(name.as_str())));
            done_jumps.push(self.emit_jump(Op::BR));
            self.patch_jump(no_match);
        }

        self.emit(Op::NULL);
        for jump in done_jumps {
            self.patch_jump(jump);
        }
        let _ = line;
        Ok(())
    }

    pub(super) fn emit_enum_value_to_string(&mut self, enum_type: &str, value_expr: &Expression) -> Result<(), String> {
        let Some(entries) = self.enum_entries_sorted(enum_type) else {
            self.compile_expr(value_expr)?;
            let to_str_idx = self.import("ecma:string", "String");
            self.emit_host_call(to_str_idx, 1);
            return Ok(());
        };

        let value_slot = self.define_local("__enum_tostring_value");
        self.compile_expr(value_expr)?;
        self.emit_u16(Op::LOCAL_SET, value_slot);
        self.emit(Op::DROP);

        let mut done_jumps = Vec::new();
        for (value, name) in &entries {
            self.emit_u16(Op::LOCAL_GET, value_slot);
            self.emit_const(Value::F64(*value as f64));
            self.emit(Op::DYN_EQ);
            let no_match = self.emit_jump(Op::BR_IF_FALSE);
            self.emit_const(Value::String(Arc::from(name.as_str())));
            done_jumps.push(self.emit_jump(Op::BR));
            self.patch_jump(no_match);
        }

        if self.enum_flags.contains(enum_type) {
            let result_slot = self.define_local("__enum_tostring_result");
            let matched_slot = self.define_local("__enum_tostring_matched");
            self.emit_const(Value::String(Arc::from("")));
            self.emit_u16(Op::LOCAL_SET, result_slot);
            self.emit(Op::DROP);
            self.emit(Op::FALSE);
            self.emit_u16(Op::LOCAL_SET, matched_slot);
            self.emit(Op::DROP);

            for (value, name) in &entries {
                if *value <= 0 || (value & (value - 1)) != 0 {
                    continue;
                }
                self.emit_u16(Op::LOCAL_GET, value_slot);
                self.emit_const(Value::F64(*value as f64));
                self.emit(Op::I32_AND);
                self.emit_const(Value::F64(*value as f64));
                self.emit(Op::DYN_EQ);
                let skip = self.emit_jump(Op::BR_IF_FALSE);

                self.emit_u16(Op::LOCAL_GET, matched_slot);
                let first = self.emit_jump(Op::BR_IF_FALSE);
                self.emit_u16(Op::LOCAL_GET, result_slot);
                self.emit_const(Value::String(Arc::from(", ")));
                self.emit(Op::DYN_ADD);
                self.emit_const(Value::String(Arc::from(name.as_str())));
                self.emit(Op::DYN_ADD);
                let with_separator = self.emit_jump(Op::BR);

                self.patch_jump(first);
                self.emit_const(Value::String(Arc::from(name.as_str())));
                self.patch_jump(with_separator);

                self.emit_u16(Op::LOCAL_SET, result_slot);
                self.emit(Op::DROP);
                self.emit(Op::TRUE);
                self.emit_u16(Op::LOCAL_SET, matched_slot);
                self.emit(Op::DROP);
                self.patch_jump(skip);
            }

            self.emit_u16(Op::LOCAL_GET, matched_slot);
            let no_flags_match = self.emit_jump(Op::BR_IF_FALSE);
            self.emit_u16(Op::LOCAL_GET, result_slot);
            done_jumps.push(self.emit_jump(Op::BR));
            self.patch_jump(no_flags_match);
        }

        self.emit_u16(Op::LOCAL_GET, value_slot);
        let to_str_idx = self.import("ecma:string", "String");
        self.emit_host_call(to_str_idx, 1);
        for jump in done_jumps {
            self.patch_jump(jump);
        }
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
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit(Op::REF_TYPEOF);
        self.emit_const(Value::String(Arc::from("number")));
        self.emit(Op::DYN_EQ);
        let not_number = self.emit_jump(Op::BR_IF_FALSE);

        let helper = self.str_const("__vybe_dotnet_numeric_format");
        self.emit_u16(Op::GLOBAL_GET, helper);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit_const(Value::String(Arc::from("F12")));
        self.emit_const(Value::F64(0.0));
        self.emit_u8(Op::CALL_REF, 3);
        let parse_float = self.import("ecma:number", "parseFloat");
        self.emit_host_call(parse_float, 1);
        let done = self.emit_jump(Op::BR);

        self.patch_jump(not_number);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.patch_jump(done);
        Ok(())
    }

    fn emit_enum_has_flag(&mut self, value_expr: &Expression, flag_expr: &Expression) -> Result<(), String> {
        let flag_slot = self.define_local("__enum_flag_value");
        let value_slot = self.define_local("__enum_flag_source");
        self.compile_expr(flag_expr)?;
        self.emit_u16(Op::LOCAL_SET, flag_slot);
        self.emit(Op::DROP);
        self.compile_expr(value_expr)?;
        self.emit_u16(Op::LOCAL_SET, value_slot);
        self.emit(Op::DROP);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit_u16(Op::LOCAL_GET, flag_slot);
        self.emit(Op::I32_AND);
        self.emit_u16(Op::LOCAL_GET, flag_slot);
        self.emit(Op::DYN_EQ);
        Ok(())
    }

    fn try_compile_dotnet_enum_call(&mut self, callee: &Expression, args: &[Argument]) -> Result<bool, String> {
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
                if receiver.rsplit('.').next().is_some_and(|type_name| type_name.eq_ignore_ascii_case("Enum")) {
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
                    let Some(enum_type) = self.canonical_enum_type_from_runtime_type(&args[0].value) else {
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
                    let Some(enum_type) = self.canonical_enum_type_from_runtime_type(&args[0].value) else {
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
                    let Some(enum_type) = self.canonical_enum_type_from_runtime_type(&args[0].value) else {
                        return Ok(false);
                    };
                    self.emit_enum_name_lookup(&enum_type, &args[1].value, false)?;
                    return Ok(true);
                }
                "IsDefined" if args.len() >= 2 => {
                    let Some(enum_type) = self.canonical_enum_type_from_runtime_type(&args[0].value) else {
                        return Ok(false);
                    };
                    self.emit_enum_name_lookup(&enum_type, &args[1].value, false)?;
                    self.emit(Op::REF_IS_NULL);
                    self.emit(Op::DYN_NOT);
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
                    let visible_args = if args.len() >= 4 { &args[..args.len() - 2] } else { args };
                    let enum_type = extract_generic_type_name(field)
                        .map(|name| self.canon(&name))
                        .filter(|canon| self.enum_value_names.contains_key(canon))
                        .or_else(|| {
                            (args.len() >= 4)
                                .then(|| self.canonical_enum_type_from_expr(&args[args.len() - 2].value))
                                .flatten()
                        });
                    let Some(enum_type) = enum_type else {
                        return Ok(false);
                    };
                    let (value_arg, ignore_case, out_arg) = if visible_args.len() == 3 {
                        (&visible_args[0].value, matches!(visible_args[1].value.kind, ExprKind::Lit(Literal::Bool(true))), &visible_args[2].value)
                    } else {
                        (&visible_args[0].value, false, &visible_args[1].value)
                    };
                    self.emit_enum_name_lookup(&enum_type, value_arg, ignore_case)?;
                    let parsed_slot = self.define_local("__enum_try_parse_value");
                    self.emit_u16(Op::LOCAL_SET, parsed_slot);
                    self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_GET, parsed_slot);
                    self.emit(Op::REF_IS_NULL);
                    let invalid = self.emit_jump(Op::BR_IF_TRUE);
                    self.emit_u16(Op::LOCAL_GET, parsed_slot);
                    self.compile_assign_target(out_arg)?;
                    self.emit(Op::TRUE);
                    let done = self.emit_jump(Op::BR);
                    self.patch_jump(invalid);
                    self.emit(Op::NULL);
                    self.compile_assign_target(out_arg)?;
                    self.emit(Op::FALSE);
                    self.patch_jump(done);
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

    pub(super) fn compile_lambda(&mut self, params: &[Param], body: &LambdaBody, captures: &[String]) -> Result<(), String> {
        let normalized_captures: Vec<String> = captures
            .iter()
            .map(|capture| self.normalize_explicit_capture(capture))
            .collect();

        if normalized_captures
            .iter()
            .any(|capture| !Self::split_explicit_capture(capture).0)
        {
            return self.compile_lambda_with_explicit_captures(params, body, &normalized_captures);
        }

        self.compile_lambda_direct(params, body)
    }

    fn compile_lambda_with_explicit_captures(
        &mut self,
        params: &[Param],
        body: &LambdaBody,
        captures: &[String],
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
        let factory = common::functions::create_function_chunk("<lambda_factory>", capture_bindings.len() as u8);
        self.chunks.push(factory);
        self.scopes.push(Scope::new_function());
        let saved = self.current;
        self.current = factory_idx;

        for (capture_name, capture_type) in &capture_bindings {
            self.define_local_typed(capture_name, capture_type.clone());
        }

        self.compile_lambda_direct(params, body)?;
        self.emit(Op::RETURN);

        let locals = self.scope().next_slot.max(self.chunks[factory_idx].local_count);
        self.chunks[factory_idx].local_count = locals;
        let uvs = self.scopes.last().unwrap().upvalues.clone();
        self.scopes.pop();
        self.current = saved;

        let line = self.line;
        common::functions::emit_ref_func(&mut self.chunks[self.current], factory_idx, uvs.len() as u8, line);
        for uv in &uvs {
            self.chunks[self.current].emit(if uv.is_local { 1 } else { 0 }, line);
            self.chunks[self.current].emit(uv.index, line);
        }
        for capture in captures {
            let (by_ref, capture_name) = Self::split_explicit_capture(capture);
            if !by_ref {
                self.emit_var_get(capture_name);
            }
        }
        self.emit_u8(Op::CALL_REF, capture_bindings.len() as u8);
        Ok(())
    }

    fn compile_lambda_direct(&mut self, params: &[Param], body: &LambdaBody) -> Result<(), String> {
        let has_rest = params.last().map_or(false, |p| p.is_rest);
        if has_rest {
            self.rest_fixed_arities.insert(params.len().saturating_sub(1) as u8);
        }
        let arity = params.len() as u8;
        let ci = self.chunks.len();
        let chunk = common::functions::create_function_chunk("<lambda>", arity);
        self.chunks.push(chunk);
        self.scopes.push(Scope::new_function());
        let saved = self.current;
        self.current = ci;
        for p in params {
            self.define_local_typed(&p.name, p.type_hint.clone());
            if let Some(ref default) = p.default {
                let slot = self.scope().resolve(&p.name).unwrap();
                self.emit_u16(Op::LOCAL_GET, slot);
                self.emit(Op::REF_IS_NULL);
                let has_val = self.emit_jump(Op::BR_IF_FALSE);
                self.compile_expr(default)?;
                self.emit_u16(Op::LOCAL_SET, slot); self.emit(Op::DROP);
                self.patch_jump(has_val);
            }
        }
        let result_slot = if self.profile.function_return == ReturnStyle::ResultSlot {
            let rs = self.define_local("Result");
            self.emit(Op::NULL); self.emit_u16(Op::LOCAL_SET, rs); self.emit(Op::DROP);
            let saved_fn = self.current_func_name.take();
            let saved_rs = self.current_result_slot.take();
            self.current_func_name = Some("<lambda>".into());
            self.current_result_slot = Some(rs);
            Some((rs, saved_fn, saved_rs))
        } else { None };

        match body {
            LambdaBody::Expr(expr) => {
                self.compile_expr(expr)?;
                self.emit(Op::RETURN);
            }
            LambdaBody::Block(stmts) => {
                for s in stmts { self.compile_stmt(s)?; }
            }
        }

        if let Some((rs, saved_fn, saved_rs)) = result_slot {
            self.emit_u16(Op::LOCAL_GET, rs);
            self.emit(Op::RETURN);
            self.current_func_name = saved_fn;
            self.current_result_slot = saved_rs;
        } else if matches!(body, LambdaBody::Block(_)) {
            let line = self.line;
            common::functions::emit_function_epilogue(&mut self.chunks[ci], line);
        }

        let locals = self.scope().next_slot.max(self.chunks[ci].local_count);
        self.chunks[ci].local_count = locals;
        let uvs = self.scopes.last().unwrap().upvalues.clone();
        self.scopes.pop();
        self.current = saved;
        let parent_locals = self.scope().locals.clone();
        let mut emitted_uvs = Vec::with_capacity(uvs.len());
        for uv in &uvs {
            let mut emitted = uv.clone();
            if uv.is_local {
                if let Some(local) = parent_locals.iter().find(|local| local.slot == uv.index as u16) {
                    let local_name = self.canon(&local.name);
                    if self.capture_by_value_vars.iter().any(|name| name == &local_name) {
                        self.emit_u16(Op::LOCAL_GET, local.slot);
                        let snapshot_slot = self.define_local_typed(
                            &format!("__capture_{}_{}", local.name, ci),
                            local.type_hint.clone(),
                        );
                        self.emit_u16(Op::LOCAL_SET, snapshot_slot);
                        self.emit(Op::DROP);
                        emitted = UpvalueDesc {
                            index: snapshot_slot as u8,
                            is_local: true,
                        };
                    }
                }
            }
            emitted_uvs.push(emitted);
        }
        let line = self.line;
        common::functions::emit_ref_func(&mut self.chunks[self.current], ci, emitted_uvs.len() as u8, line);
        for uv in &emitted_uvs {
            self.chunks[self.current].emit(if uv.is_local { 1 } else { 0 }, line);
            self.chunks[self.current].emit(uv.index, line);
        }
        if has_rest {
            self.emit_stamp_rest_metadata_on_stack(params.len().saturating_sub(1));
        }
        Ok(())
    }

}
