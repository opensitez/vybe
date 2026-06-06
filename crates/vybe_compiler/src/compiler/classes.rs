//! Class, constructor, and free-function compilation.
//!
//! Extracted from `compiler.rs` to keep that file navigable. The
//! methods on this `impl Compiler { ... }` block are private by
//! convention (they're only called from other compiler methods) and
//! crate-private for the `dotnet_register` bridge.

use super::*;
use crate::common::classes::{BaseCall, NormalConstructor, NormalMethod};
use crate::compiler::ArrayBindingMetadata;
use crate::scope::UpvalueDesc;

impl Compiler {
    fn fixed_array_zero_expr(type_hint: &str) -> Option<Expression> {
        let trimmed = type_hint.trim();
        if !trimmed.starts_with('[') {
            return None;
        }

        let close = trimmed.find(']')?;
        let head = trimmed.get(1..close)?.trim();
        if head.is_empty() || head == "..." {
            return None;
        }

        let len = head.parse::<usize>().ok()?;
        let element_type = trimmed.get(close + 1..)?.trim();
        let element_expr = Self::fixed_array_zero_expr(element_type)
            .unwrap_or_else(|| Self::array_default_element_expr(Some(element_type)));

        Some(Expression::new(ExprKind::Cast {
            expr: Box::new(Expression::new(ExprKind::Array(
                (0..len)
                    .map(|_| ArrayElement {
                        key: None,
                        value: element_expr.clone(),
                        spread: false,
                        by_ref: false,
                    })
                    .collect(),
            ))),
            type_name: trimmed.to_string(),
        }))
    }

    fn array_default_element_expr(type_hint: Option<&str>) -> Expression {
        match type_hint
            .map(str::trim)
            .map(|hint| hint.strip_suffix("()").unwrap_or(hint))
            .map(Self::normalize_type_hint)
            .as_deref()
        {
            Some("integer") | Some("int") | Some("int32") | Some("longint") | Some("real")
            | Some("double") | Some("float") | Some("single") | Some("decimal") | Some("long")
            | Some("int64") | Some("short") | Some("int16") | Some("uint") | Some("uint32")
            | Some("ulong") | Some("uint64") | Some("ushort") | Some("uint16") | Some("byte")
            | Some("sbyte") | Some("char") => Expression::new(ExprKind::Lit(Literal::Int(0))),
            Some("boolean") | Some("bool") => Expression::new(ExprKind::Lit(Literal::Bool(false))),
            Some(type_hint) if Self::is_string_type_hint(type_hint) => Expression::string(""),
            _ => Expression::null(),
        }
    }

    fn array_bounds_extent_expr(bounds: &[Expression]) -> Option<Expression> {
        let mut iter = bounds.iter().cloned();
        let first = iter.next()?;
        Some(iter.fold(first, |acc, bound| {
            Expression::new(ExprKind::Binary {
                op: BinOp::Mul,
                left: Box::new(acc),
                right: Box::new(bound),
            })
        }))
    }

    fn emit_class_field_initializer(
        &mut self,
        owner_slot: u16,
        field_name: &str,
        type_hint: Option<&str>,
        init: Option<&Expression>,
        array_bounds: Option<&[Expression]>,
        is_value_type: bool,
        line: u32,
    ) -> Result<(), String> {
        let value_slot = self.define_local("__field_init_value");
        if let Some(init_expr) = init {
            self.compile_expr(init_expr)?;
        } else if let Some(extent) = array_bounds.and_then(Self::array_bounds_extent_expr) {
            if let Some(init_expr) = type_hint.and_then(Self::fixed_array_zero_expr) {
                self.compile_expr(&init_expr)?;
            } else {
                let init_expr = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("Array")),
                    args: vec![
                        Argument::positional(extent),
                        Argument::positional(Self::array_default_element_expr(type_hint)),
                    ],
                    optional: false,
                });
                self.compile_expr(&init_expr)?;
            }
        } else if let Some(init_expr) = type_hint.and_then(Self::fixed_array_zero_expr) {
            self.compile_expr(&init_expr)?;
        } else if let Some(type_name) =
            type_hint.and_then(|type_hint| self.user_value_type_name_from_hint(type_hint))
        {
            let ctor_global = {
                let overload = format!("{}$arity0", type_name);
                if self.defined_globals.contains(&overload) {
                    overload
                } else {
                    type_name.clone()
                }
            };
            let idx = self.str_const(&ctor_global);
            self.emit_u16(Op::GLOBAL_GET, idx);
            self.emit_u8(Op::CALL_REF, 0);
        } else if is_value_type {
            self.emit_default_value_for_type_hint(type_hint);
        } else if self.is_js_profile() {
            // JS spec: declared fields with no initializer default to undefined (not null).
            self.emit(Op::UNDEFINED);
        } else {
            self.emit(Op::NULL);
        }
        self.emit_u16(Op::LOCAL_SET, value_slot);
        self.emit(Op::DROP);

        common::classes::emit_init_field_start(self.chunk(), owner_slot, line);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        common::classes::emit_init_field_end(self.chunk(), field_name, line);
        Ok(())
    }

    fn class_requires_form_identity_stamp(&self, parent: &Option<String>) -> bool {
        let mut current = parent.clone().map(|name| self.canon(&name));
        let mut visited = std::collections::HashSet::new();

        while let Some(name) = current {
            if !visited.insert(name.clone()) {
                break;
            }
            if name.eq_ignore_ascii_case("form")
                || self.reflection_is_assignable_from("Form", &name)
            {
                return true;
            }
            current = self
                .pending_classes
                .get(name.as_str())
                .and_then(|pending| pending.parent.clone())
                .or_else(|| self.reflection_base_type_name(&name));
        }

        false
    }

    fn emit_form_identity_stamp(&mut self, this_slot: u16, class_name: &str, _line: u32) {
        let stamped_name = self.canon(class_name);
        let set_property = self.import("vybe:gui", "controlSetProperty");
        self.emit_u16(Op::LOCAL_GET, this_slot);
        self.emit_const(Value::String(Arc::from("Name")));
        self.emit_const(Value::String(Arc::from(stamped_name.as_str())));
        self.emit_host_call(set_property, 3);
        self.emit(Op::DROP);
    }

    fn emit_store_super_ref(&mut self, this_slot: u16, parent_name: &str) {
        self.emit_u16(Op::LOCAL_GET, this_slot);
        self.emit_var_get(parent_name);
        let super_key = self.str_const("__super");
        self.emit_u16(Op::STRUCT_SET, super_key);
        self.emit(Op::DROP);
    }

    fn captured_name_for_upvalue(&self, scope_idx: usize, upvalue_idx: u8) -> Option<String> {
        let upvalue = self
            .scopes
            .get(scope_idx)?
            .upvalues
            .get(upvalue_idx as usize)?;
        let parent_scope_idx = scope_idx.checked_sub(1)?;
        if upvalue.is_local {
            self.scopes
                .get(parent_scope_idx)?
                .locals
                .iter()
                .find(|local| local.slot == upvalue.index as u16)
                .map(|local| local.name.clone())
        } else {
            self.captured_name_for_upvalue(parent_scope_idx, upvalue.index)
        }
    }

    fn emit_ref_func_with_captures(
        &mut self,
        func_idx: usize,
        capture_names: &[String],
    ) -> Result<(), String> {
        let line = self.line;
        common::functions::emit_ref_func(
            &mut self.chunks[self.current],
            func_idx,
            capture_names.len() as u8,
            line,
        );
        for capture_name in capture_names {
            if let Some(slot) = self
                .scope()
                .resolve(capture_name)
                .or_else(|| self.scope().resolve_ci(capture_name))
            {
                self.chunks[self.current].emit(1, line);
                self.chunks[self.current].emit(slot as u8, line);
                continue;
            }
            if self.scopes.len() > 1 {
                if let Some(uv) = self.resolve_upvalue(self.scopes.len() - 1, capture_name) {
                    self.chunks[self.current].emit(0, line);
                    self.chunks[self.current].emit(uv, line);
                    continue;
                }
            }
            return Err(format!(
                "failed to resolve captured class method binding '{capture_name}'"
            ));
        }
        Ok(())
    }

    fn emit_bind_instance_method_with_aliases(
        &mut self,
        this_slot: u16,
        method_name: &str,
        method_chunk_idx: usize,
        capture_names: &[String],
        rest_fixed_count: Option<u8>,
        bind_receiver: bool,
    ) -> Result<(), String> {
        let receiver_key = self.str_const("__vybe_method_receiver");
        let rest_key = self.str_const("__vybe_rest_fixed_arity");

        let mut bind_names = vec![method_name.to_string()];
        for &alias in common::classes::cross_language_aliases(method_name) {
            if alias != method_name {
                bind_names.push(alias.to_string());
            }
        }

        for bind_name in bind_names {
            self.emit_u16(Op::LOCAL_GET, this_slot);
            self.emit_ref_func_with_captures(method_chunk_idx, capture_names)?;
            if bind_receiver {
                self.emit(Op::DUP);
                self.emit_u16(Op::LOCAL_GET, this_slot);
                self.emit_u16(Op::STRUCT_SET, receiver_key);
                self.emit(Op::DROP);
            }
            if let Some(fixed_count) = rest_fixed_count {
                self.emit(Op::DUP);
                self.emit_const(Value::F64(fixed_count as f64));
                self.emit_u16(Op::STRUCT_SET, rest_key);
                self.emit(Op::DROP);
            }
            let method_key = self.str_const(&bind_name);
            self.emit_u16(Op::STRUCT_SET, method_key);
            self.emit(Op::DROP);
        }
        Ok(())
    }

    // ════════════════════════════════════════════════════════════════════════
    // Function declaration compilation
    // ════════════════════════════════════════════════════════════════════════

    pub(super) fn compile_function_decl(
        &mut self,
        name: &str,
        params: &[Param],
        return_type: &Option<String>,
        body: &[Statement],
        _is_sub: bool,
        is_generator: bool,
        handles: &[String],
        is_async: bool,
    ) -> Result<(), String> {
        let cname = self.canon(name);
        self.defined_globals.insert(cname.clone());
        self.defined_functions.insert(cname.clone());
        // Register top-level generator functions so `is_direct_generator_call`
        // detects them (`[...gen()]` spread, `foreach (gen() as ...)`). Scoped
        // to buffered-iterator languages (PHP): JS keeps its runtime
        // `isGenerator` dispatch, so registering here would change its routing.
        if is_generator && self.profile.buffered_iterator_methods {
            self.generator_functions.insert(cname.clone());
        }
        self.function_param_modes.insert(
            cname.clone(),
            params.iter().map(|param| param.pass_by).collect(),
        );
        self.function_param_types.insert(
            cname.clone(),
            params.iter().map(|param| param.type_hint.clone()).collect(),
        );
        self.function_min_arity.insert(
            cname.clone(),
            params
                .iter()
                .take_while(|param| param.default.is_none() && !param.is_rest)
                .count(),
        );
        self.function_signatures
            .entry(cname.clone())
            .or_default()
            .push(CallSignature::from_params(params));
        if let Some(return_type) = return_type.as_ref() {
            self.function_return_types
                .insert(cname.clone(), return_type.clone());
        }
        let name = &cname;

        let uses_js_arguments = self.is_js_profile()
            && !is_generator
            && (params
                .iter()
                .any(|param| param.default.as_ref().is_some_and(expr_uses_js_arguments))
                || body.iter().any(stmt_uses_js_arguments));
        let has_rest = params.last().map_or(false, |p| p.is_rest);
        let lowered_has_rest = has_rest || uses_js_arguments;
        let generator_control_arity = usize::from(is_generator && !lowered_has_rest);
        let arity: u8 = if uses_js_arguments {
            (1 + generator_control_arity) as u8
        } else {
            (params.len() + generator_control_arity) as u8
        };
        if uses_js_arguments {
            self.rest_fixed_arities.insert(0);
        } else if has_rest {
            self.rest_fixed_arities
                .insert(params.len().saturating_sub(1) as u8);
        }
        let func_idx = self.chunks.len();
        let mut chunk = common::functions::create_function_chunk(name, arity);
        self.seed_shared_global_constants(&mut chunk);
        // Mark the chunk so the WASM emitter can list it in the
        // `vybe.jspi` custom section — JS hosts wrap promising exports
        // via `WebAssembly.promising(...)` at load time.
        chunk.is_async = is_async;
        // Generators: when the source marked the function as a
        // generator (Python `yield`, JS `function*`, C# `yield return`),
        // stamp the chunk so the VM wraps invocations in a
        // `Continuation` instead of executing the body inline. The
        // body itself was compiled with `SUSPEND` at each yield site.
        chunk.is_generator = is_generator;
        if is_generator && !is_async {
            self.generator_functions.insert(cname.clone());
            // Track the number of user-visible params so call sites can pad
            // missing optional args with `undefined`.  Without this, a call
            // like `range(1, 6)` to `function* range(start, end, step=1)`
            // passes bound_args=[1,6]; GEN_NEXT then calls the body with
            // [1, 6, null] (argc=3) which lands `null` in the `step` slot,
            // preventing the default from applying and causing an infinite loop.
            // Only register fixed-arity generators (no rest, no arguments).
            if !lowered_has_rest {
                self.generator_param_counts
                    .insert(cname.clone(), params.len());
            }
        }
        // Multi-value tuple returns: if the pre-scan marked this function
        // as a same-arity multi-tuple-return, stamp its result_arity here
        // so the WASM type section emits `(externref^N) -> (externref^N)`
        // and the VM's RETURN knows to pop N values off the stack.
        if let Some(&n) = self.multi_return_functions.get(&cname) {
            chunk.result_arity = n;
        }
        self.chunks.push(chunk);
        self.scopes.push(Scope::new_function());
        self.static_local_bindings.push(HashMap::new());
        let saved = self.current;
        self.current = func_idx;
        // Function body opens fresh wrt the runtime label_stack —
        // emit_return drains back to this base. Save+restore so nested
        // function decls compose.
        let saved_label_base = self.function_label_base;
        self.function_label_base = self.label_depth;
        let saved_fn = self.current_func_name.take();
        self.current_func_name = Some(name.to_string());
        self.js_arguments_bindings.push(None);

        let js_arguments_source_slot = if uses_js_arguments {
            Some(self.define_local("__vybe_js_arguments_array"))
        } else {
            None
        };
        let js_arguments_slot = if uses_js_arguments {
            let slot = self.define_local("arguments");
            self.emit_u16(Op::LOCAL_GET, js_arguments_source_slot.unwrap());
            self.emit_u16(Op::LOCAL_SET, slot);
            self.emit(Op::DROP);
            self.emit_u16(Op::LOCAL_GET, slot);
            self.emit_var_get(name);
            let callee_key = self.str_const("callee");
            self.emit_u16(Op::STRUCT_SET, callee_key);
            self.emit(Op::DROP);
            Some(slot)
        } else {
            None
        };

        let mut aliased_params = HashMap::new();
        let mut aliased_indices = HashMap::new();
        let simple_arguments_alias = uses_js_arguments
            && params
                .iter()
                .all(|param| param.default.is_none() && !param.is_rest);

        for (index, p) in params.iter().enumerate() {
            self.define_local_typed(&p.name, p.type_hint.clone());
            let normalized_type_hint = p.type_hint.as_deref().map(Compiler::normalize_type_hint);
            if normalized_type_hint
                .as_deref()
                .is_some_and(|type_hint| type_hint.ends_with("()"))
                || normalized_type_hint
                    .as_deref()
                    .is_some_and(|type_hint| type_hint.ends_with("()"))
            {
                self.record_array_binding(
                    &p.name,
                    ArrayBindingMetadata {
                        is_fixed: false,
                        type_hint: p.type_hint.clone(),
                        pascal_bounds: p
                            .type_hint
                            .as_deref()
                            .and_then(|type_hint| self.pascal_array_type_hint_metadata(type_hint)),
                    },
                );
            }
            if simple_arguments_alias {
                let slot = self.scope().resolve(&p.name).unwrap();
                aliased_params.insert(p.name.clone(), (slot, index));
                aliased_indices.insert(index, slot);
            }
        }

        let js_arguments_len_slot = if uses_js_arguments {
            let len_slot = self.define_local("__vybe_js_arguments_length");
            self.emit_u16(Op::LOCAL_GET, js_arguments_source_slot.unwrap());
            common::collections::emit_len(&mut self.chunks, self.current, self.line);
            self.emit_u16(Op::LOCAL_SET, len_slot);
            self.emit(Op::DROP);
            Some(len_slot)
        } else {
            None
        };

        if let Some(slot) = js_arguments_slot {
            *self.js_arguments_bindings.last_mut().unwrap() = Some(JsArgumentsBinding {
                args_slot: slot,
                aliased_params,
                aliased_indices,
            });
        }

        if uses_js_arguments {
            for (index, p) in params.iter().enumerate() {
                let slot = self.scope().resolve(&p.name).unwrap();
                if p.is_rest {
                    self.emit_u16(Op::LOCAL_GET, js_arguments_source_slot.unwrap());
                    self.emit_const(Value::F64(index as f64));
                    self.emit_u16(Op::LOCAL_GET, js_arguments_len_slot.unwrap());
                    common::collections::emit_slice(&mut self.chunks, self.current, self.line);
                    self.emit_u16(Op::LOCAL_SET, slot);
                    self.emit(Op::DROP);
                } else {
                    self.emit_array_value_or_undefined(
                        js_arguments_source_slot.unwrap(),
                        js_arguments_len_slot.unwrap(),
                        index,
                    );
                    self.emit_u16(Op::LOCAL_SET, slot);
                    self.emit(Op::DROP);
                }
            }
        }

        for p in params {
            // Default parameters: ECMA-262 §15.2.3 — only `undefined`
            // triggers the default (not `null`). The VM now pads
            // missing positional args with `Undefined`, distinct from
            // an explicitly-passed `Null`, so `REF_IS_UNDEFINED` is
            // the correct discriminant for JS. Other languages don't
            // distinguish missing/null and use `REF_IS_NULL` (matches
            // either tag).
            if let Some(ref default) = p.default {
                let slot = self.scope().resolve(&p.name).unwrap();
                self.emit_u16(Op::LOCAL_GET, slot);
                if self.is_js_profile() {
                    self.emit(Op::REF_IS_UNDEFINED);
                } else {
                    self.emit(Op::REF_IS_NULL);
                }
                let branch_line = self.line;
                crate::emitter::ops::emit_dyn_to_bool(self.chunk(), branch_line);
                self.chunks[self.current].emit_if(branch_line);
                self.compile_expr(default)?;
                self.emit_u16(Op::LOCAL_SET, slot);
                self.emit(Op::DROP);
                self.chunks[self.current].emit_end(branch_line);
            }
            self.maybe_initialize_fortran_out_param(p);
        }

        let generator_control_slot =
            is_generator.then(|| self.define_local("__generator_entry_control"));

        if let Some(control_slot) = generator_control_slot {
            self.emit_generator_entry_control(control_slot)?;
        }

        // Result slot for functions with return type (Pascal/VB Function).
        // The slot name is profile-driven so VB can keep it internal
        // (`__result__`) and avoid shadowing user classes named `Result`,
        // while Pascal keeps it as `Result` (user-visible per Pascal idiom).
        let result_slot =
            if return_type.is_some() && self.profile.function_return == ReturnStyle::ResultSlot {
                let slot_name = self.profile.result_slot_name.clone();
                let rs = self.define_local(&slot_name);
                self.emit(Op::NULL);
                self.emit_u16(Op::LOCAL_SET, rs);
                self.emit(Op::DROP);
                Some(rs)
            } else {
                None
            };
        let ref_out_slots: Vec<u16> = params
            .iter()
            .filter(|param| matches!(param.pass_by, PassBy::Ref | PassBy::Out))
            .filter_map(|param| self.scope().resolve(&param.name))
            .collect();

        let saved_rs = self.current_result_slot.take();
        let saved_ref_out = self.current_ref_out_params.take();
        self.current_result_slot = result_slot;
        self.current_ref_out_params = (!ref_out_slots.is_empty()).then_some(ref_out_slots);

        // ECMA-262 async function semantics: throws inside the body
        // become rejected Promises, normal returns become fulfilled
        // Promises (§27.7.5.3). Wrap the body in TRY_START/TRY_END so
        // uncaught exceptions short-circuit to the Promise.reject
        // path. The `await` opcode handles per-await suspension via
        // JSPI; this wrap just covers terminal throw / return.
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

        if self.profile.name == "fortran" {
            for statement in body {
                if matches!(&statement.kind, StmtKind::VarDecl { .. }) {
                    self.compile_stmt(statement)?;
                }
            }
            for statement in body {
                if matches!(&statement.kind, StmtKind::FunctionDecl { .. }) {
                    self.compile_stmt(statement)?;
                }
            }
            for statement in body {
                if matches!(
                    &statement.kind,
                    StmtKind::VarDecl { .. } | StmtKind::FunctionDecl { .. }
                ) {
                    continue;
                }
                self.compile_stmt(statement)?;
            }
        } else {
            for statement in body {
                self.compile_stmt(statement)?;
            }
        }

        if async_try.is_some() {
            self.active_async_try_depth = self.active_async_try_depth.saturating_sub(1);
        }

        if let Some(catch_jump) = async_try {
            let line = self.line;
            // Normal exit: wrap return in Promise.resolve(value).
            // The body's compile_stmt may have already emitted RETURNs
            // (early returns); we still need the fall-through path
            // to leave a fulfilled Promise on the stack.
            let chunk = &mut self.chunks[self.current];
            common::functions::emit_async_body_fallthrough(chunk, catch_jump, line);
            let resolve_idx = self.import("ecma:promise", "resolve");
            self.emit_host_call(resolve_idx, 1);
            self.emit_return();
            // Catch handler — exception value on TOS.
            let chunk = &mut self.chunks[self.current];
            common::functions::patch_async_body_catch(chunk, catch_jump);
            let reject_idx = self.import("ecma:promise", "reject");
            self.emit_host_call(reject_idx, 1);
            self.emit_return();
        } else if let Some(rs) = result_slot {
            self.emit_u16(Op::LOCAL_GET, rs);
            self.emit_return_through_finally(1)?;
        } else if self.current_ref_out_params.is_some() {
            self.emit(Op::NULL);
            self.emit_return_through_finally(1)?;
        } else {
            let line = self.line;
            common::functions::emit_function_epilogue(&mut self.chunks[func_idx], line);
        }

        self.current_func_name = saved_fn;
        self.current_result_slot = saved_rs;
        self.current_ref_out_params = saved_ref_out;

        let locals = self
            .scope()
            .next_slot
            .max(self.chunks[func_idx].local_count);
        self.chunks[func_idx].local_count = locals;
        let uvs = self.scopes.last().unwrap().upvalues.clone();
        self.js_arguments_bindings.pop();
        self.scopes.pop();
        self.static_local_bindings.pop();
        self.current = saved;
        self.function_label_base = saved_label_base;

        let line = self.line;
        common::functions::emit_ref_func(
            &mut self.chunks[self.current],
            func_idx,
            uvs.len() as u8,
            line,
        );
        for uv in &uvs {
            self.chunks[self.current].emit(if uv.is_local { 1 } else { 0 }, line);
            self.chunks[self.current].emit(uv.index, line);
        }
        if uses_js_arguments {
            self.emit_stamp_rest_metadata_on_stack(0);
        } else if has_rest {
            self.emit_stamp_rest_metadata_on_stack(params.len().saturating_sub(1));
        }
        let idx = self.str_const(name);
        self.emit_u16(Op::GLOBAL_SET, idx);
        self.emit(Op::DROP);

        if self.is_js_profile() {
            let line = self.line;
            self.emit_common("object.new", 0, line);
            let proto_slot = self.define_local("__js_fn_proto");
            self.emit_u16(Op::LOCAL_SET, proto_slot);
            self.emit(Op::DROP);

            self.emit_var_get(name);
            self.emit_const(Value::String(Arc::from(name.as_str())));
            let name_key = self.str_const("name");
            self.emit_u16(Op::STRUCT_SET, name_key);
            self.emit(Op::DROP);

            // ECMA-262 §10.2.4 `length`: number of params before the
            // first one with a default value or rest. Skip rest entirely.
            let length = params
                .iter()
                .take_while(|p| p.default.is_none() && !p.is_rest)
                .count();
            self.emit_var_get(name);
            self.emit_const(Value::F64(length as f64));
            let length_key = self.str_const("length");
            self.emit_u16(Op::STRUCT_SET, length_key);
            self.emit(Op::DROP);

            self.emit_var_get(name);
            self.emit_var_get("Function");
            let function_proto_key = self.str_const("prototype");
            self.emit_u16(Op::STRUCT_GET, function_proto_key);
            let proto_link_key = self.str_const("__proto__");
            self.emit_u16(Op::STRUCT_SET, proto_link_key);
            self.emit(Op::DROP);

            self.emit_u16(Op::LOCAL_GET, proto_slot);
            self.emit_var_get(name);
            let ctor_key = self.str_const("constructor");
            self.emit_u16(Op::STRUCT_SET, ctor_key);
            self.emit(Op::DROP);

            self.emit_var_get(name);
            self.emit_u16(Op::LOCAL_GET, proto_slot);
            let proto_key = self.str_const("prototype");
            self.emit_u16(Op::STRUCT_SET, proto_key);
            self.emit(Op::DROP);
        }

        // VB `Handles ctrl.Event` clause on a top-level Sub: register the
        // event handler with the canonical GUI binding. The same canonical
        // emit path serves C# `+=`, JS `addEventListener`, etc.
        for handle in handles {
            let parts: Vec<&str> = handle.splitn(2, '.').collect();
            if parts.len() == 2 {
                let line = self.line;
                let bind_idx = self.import("vybe:gui", common::gui::HOST_FN_BIND_EVENT);
                let ctrl_raw = parts[0].trim();
                let ctrl_canon = self.canon(ctrl_raw);
                let ctrl_key = if ctrl_canon == self.profile.self_keyword
                    || ctrl_canon == "me"
                    || ctrl_canon == "this"
                    || ctrl_canon == "mybase"
                {
                    self.current_class
                        .clone()
                        .map(|c| self.canon(&c))
                        .unwrap_or(ctrl_canon)
                } else {
                    ctrl_canon
                };
                self.emit_const(Value::String(Arc::from(ctrl_key.as_str())));
                let ev = parts[1].to_lowercase();
                self.emit_const(Value::String(Arc::from(ev.as_str())));
                self.emit_var_get(name);
                common::gui::emit_bind_event(self.chunk(), bind_idx, line);
                self.emit(Op::DROP); // statement: discard host call result
            }
        }

        Ok(())
    }

    // ════════════════════════════════════════════════════════════════════════
    // Class compilation
    // ════════════════════════════════════════════════════════════════════════

    pub(crate) fn compile_class(
        &mut self,
        class: &crate::common::classes::NormalClass,
    ) -> Result<(), String> {
        // Extract the canonicalised names the orchestration below needs.
        // Canonicalisation happens once here rather than at every caller.
        let cname = self.canon(&class.name);
        let name: &str = &cname;
        let parent_canonical = class.parent.as_ref().map(|p| self.canon(p));
        let parent: &Option<String> = &parent_canonical;

        // Phase 2b.2 complete: passes 1-4 all read NormalClass fields
        // directly. No more ClassMember reconstruction inside
        // compile_class.
        let self_kw = self.profile.self_keyword.clone();
        let ctor_name = self.profile.constructor_name.clone();
        let result_style = self.profile.function_return.clone();

        // Pass 1 (ported to NormalClass): collect fields + initialisers
        // from instance_fields / static_fields, then add backing fields
        // for auto-properties. Reads NormalClass directly; no longer
        // iterates the reconstructed member list.
        let mut fields: Vec<String> = Vec::new();
        let mut field_inits: Vec<(
            String,
            Option<String>,
            Option<Expression>,
            Option<Vec<Expression>>,
        )> = Vec::new();
        let mut static_field_inits: Vec<(
            String,
            Option<String>,
            Option<Expression>,
            Option<Vec<Expression>>,
        )> = Vec::new();
        for f in &class.instance_fields {
            let fname = self.js_member_storage_name_for_class(&class.name, &f.name);
            fields.push(fname.clone());
            field_inits.push((
                fname,
                f.type_hint.clone(),
                f.init.clone(),
                f.array_bounds.clone(),
            ));
        }
        for f in &class.static_fields {
            let fname = self.js_member_storage_name_for_class(&class.name, &f.name);
            static_field_inits.push((
                fname,
                f.type_hint.clone(),
                f.init.clone(),
                f.array_bounds.clone(),
            ));
        }
        for p in &class.properties {
            // Auto-properties get a backing field named like the property;
            // the runtime reads/writes through auto-emitted __get_/__set_
            // chunks bound later.
            if let Some(auto_field_name) = &p.auto_field {
                let pname_canon =
                    self.js_member_storage_name_for_class(&class.name, auto_field_name);
                if p.is_static {
                    if !static_field_inits
                        .iter()
                        .any(|(n, _, _, _)| n == &pname_canon)
                    {
                        static_field_inits.push((pname_canon, None, None, None));
                    }
                } else if !fields.contains(&pname_canon) {
                    fields.push(pname_canon.clone());
                    field_inits.push((pname_canon, None, None, None));
                }
            }
        }

        // Events need backing storage on instances so method bodies can
        // read/invoke them via implicit-self resolution (`if (Click != null)
        // Click();`) and subscriptions (`obj.Click += handler`) persist.
        for m in &class.raw_extra_members {
            if let ClassMember::Event { name: ename, .. } = m {
                let fname = self.canon(ename);
                if !fields.contains(&fname) {
                    fields.push(fname.clone());
                    field_inits.push((fname, None, None, None));
                }
            }
        }

        // Store field list for implicit self resolution
        let static_field_names: Vec<String> = static_field_inits
            .iter()
            .map(|(n, _, _, _)| n.clone())
            .collect();
        let mut instance_field_types: HashMap<String, String> = class
            .instance_fields
            .iter()
            .filter_map(|f| {
                f.type_hint.as_ref().map(|t| {
                    (
                        self.js_member_storage_name_for_class(&class.name, &f.name),
                        Self::normalize_type_hint(t),
                    )
                })
            })
            .collect();
        for member in &class.raw_extra_members {
            match member {
                ClassMember::Event {
                    name,
                    type_hint: Some(type_hint),
                    ..
                } => {
                    instance_field_types
                        .entry(self.canon(name))
                        .or_insert_with(|| Self::normalize_type_hint(type_hint));
                }
                ClassMember::Property {
                    name,
                    type_hint: Some(type_hint),
                    modifiers,
                    ..
                } if !modifiers.is_static => {
                    instance_field_types
                        .entry(self.canon(name))
                        .or_insert_with(|| Self::normalize_type_hint(type_hint));
                }
                _ => {}
            }
        }
        let mut static_member_names = static_field_names;
        let mut static_const_names: Vec<String> = Vec::new();
        for member in &class.raw_extra_members {
            if let ClassMember::Const { name, .. } = member {
                let const_name = self.canon(name);
                static_member_names.push(const_name.clone());
                static_const_names.push(const_name);
            }
        }

        self.pending_classes.insert(
            name.to_string(),
            PendingClass {
                parent: parent.clone(),
                enclosing_class: self.current_class.clone(),
                fields: fields.clone(),
                is_value_type: class.is_value_type,
                instance_member_names: class
                    .instance_methods
                    .iter()
                    .map(|method| {
                        self.js_member_storage_name_for_class(&class.name, &method.source_name)
                    })
                    .collect(),
                instance_pointer_method_names: class
                    .instance_methods
                    .iter()
                    .filter(|method| {
                        method
                            .params
                            .first()
                            .and_then(|param| param.type_hint.as_deref())
                            .is_some_and(|type_hint| type_hint.trim_start().starts_with('*'))
                    })
                    .map(|method| method.canonical_name.clone())
                    .collect(),
                instance_field_types,
                static_fields: static_member_names,
                static_field_types: class
                    .static_fields
                    .iter()
                    .filter_map(|f| {
                        f.type_hint
                            .as_ref()
                            .map(|t| (self.canon(&f.name), Self::normalize_type_hint(t)))
                    })
                    .collect(),
                static_method_names: class
                    .static_methods
                    .iter()
                    .map(|m| self.js_member_storage_name_for_class(&class.name, &m.source_name))
                    .collect(),
                instance_method_overloads: HashMap::new(),
                static_method_overloads: HashMap::new(),
                nested_types: class
                    .raw_extra_members
                    .iter()
                    .filter_map(|m| {
                        if let ClassMember::NestedType(stmt) = m {
                            match &stmt.kind {
                                StmtKind::ClassDecl { name, .. }
                                | StmtKind::StructDecl { name, .. }
                                | StmtKind::InterfaceDecl { name, .. }
                                | StmtKind::EnumDecl { name, .. } => Some(name.clone()),
                                _ => None,
                            }
                        } else {
                            None
                        }
                    })
                    .collect(),
                statics: Vec::new(), // filled after methods are compiled
            },
        );

        // Compile methods (including constructor body)
        // (name, chunk_idx, is_ctor, is_static)
        let mut method_chunks: Vec<(String, usize, bool, bool)> = Vec::new();
        let mut method_capture_name_map: HashMap<usize, Vec<String>> = HashMap::new();
        let saved_class = self.current_class.take();
        let saved_implicit = self.current_class_implicit_self;
        self.current_class = Some(name.to_string());
        self.current_class_implicit_self = class.implicit_self_fields;

        // Pass 2 (ported to NormalClass): pre-register method + property
        // names in `defined_class_methods` so expression-compilation
        // doesn't hijack a method call via the value-method dispatch
        // table. Walks instance_methods + static_methods + properties
        // directly; no reconstructed member iteration.
        for m in class
            .instance_methods
            .iter()
            .chain(class.static_methods.iter())
        {
            // Use `source_name` so existing compile paths that look up
            // `self.defined_class_methods.contains("ToString")` (from VB
            // call-site compilation) still hit. Canonical-name-only
            // lookups are a Phase 2b.3 concern.
            self.defined_class_methods
                .insert(self.canon(&m.source_name));
            if let Some(private_name) =
                self.js_private_member_storage_name_for_class(&class.name, &m.source_name)
            {
                self.defined_class_methods.insert(private_name);
            }
            if let Some(arity) = uniform_tuple_return_arity(&m.body) {
                let bound_name = if let Some(private_name) =
                    self.js_private_member_storage_name_for_class(&class.name, &m.source_name)
                {
                    private_name
                } else if m.source_name.starts_with("Symbol.") && !m.canonical_name.is_empty() {
                    m.canonical_name.clone()
                } else {
                    self.canon(&m.source_name)
                };
                self.multi_return_functions.insert(bound_name, arity);
                self.multi_return_functions.insert(
                    self.canon(&format!("{}.{}", class.name, m.source_name)),
                    arity,
                );
            }
        }
        for p in &class.properties {
            self.defined_class_methods
                .insert(self.canon(&p.source_name));
            if let Some(private_name) =
                self.js_private_member_storage_name_for_class(&class.name, &p.source_name)
            {
                self.defined_class_methods.insert(private_name);
            }
        }

        // Pass 3 (ported to NormalClass): compile method chunks,
        // property getter/setter chunks, class-level constants, and
        // nested types. Order matches the former reconstructed-
        // `members` layout so chunk indices stay byte-identical:
        //   instance_methods → static_methods → raw_extra_members
        //   → properties. Constructor body is handled in pass-4 below.

        // --- Instance + static methods ---
        // Each NormalMethod carries the walker's raw modifiers, source
        // name, params, body, return_type, is_generator, is_static
        // flag (implied by which vec the method lives in). That's all
        // the old Method arm needed.
        let mut compile_normal_method = |cc: &mut Compiler,
                                         m: &NormalMethod,
                                         is_static: bool|
         -> Result<(), String> {
            let mname = &m.source_name;
            let is_static_init = is_static && mname == "__static_init__";
            let is_ctor = if cc.case_sensitive {
                mname == &ctor_name || (is_static && mname == "new")
            } else {
                mname.eq_ignore_ascii_case(&ctor_name)
                    || is_static && mname.eq_ignore_ascii_case("new")
            };

            let user_params: Vec<&Param> = if class.explicit_self_param {
                m.params.iter().skip(1).collect()
            } else {
                m.params.iter().collect()
            };
            let param_types: Vec<String> = user_params
                .iter()
                .map(|param| {
                    Compiler::normalize_type_hint(param.type_hint.as_deref().unwrap_or("object"))
                })
                .collect();
            let bound_name = if let Some(private_name) =
                cc.js_private_member_storage_name_for_class(&class.name, mname)
            {
                private_name
            } else if mname.starts_with("Symbol.") && !m.canonical_name.is_empty() {
                m.canonical_name.clone()
            } else {
                cc.canon(mname)
            };
            let qualified_name = cc.canon(&format!("{}.{}", class.name, mname));
            cc.function_param_modes.insert(
                bound_name.clone(),
                user_params.iter().map(|param| param.pass_by).collect(),
            );
            cc.function_param_modes.insert(
                qualified_name.clone(),
                user_params.iter().map(|param| param.pass_by).collect(),
            );
            cc.function_signatures
                .entry(bound_name.clone())
                .or_default()
                .push(CallSignature::from_params(
                    &user_params
                        .iter()
                        .map(|param| (*param).clone())
                        .collect::<Vec<_>>(),
                ));
            if let Some(return_type) = m.return_type.as_ref() {
                cc.function_return_types
                    .insert(bound_name.clone(), return_type.clone());
                cc.function_return_types
                    .insert(qualified_name, return_type.clone());
                cc.function_return_types.insert(
                    cc.canon(&format!("{}.{}", class.name, mname)),
                    return_type.clone(),
                );
            }
            let uses_js_this = cc.is_js_profile();
            let has_rest = user_params.last().map_or(false, |p| p.is_rest);
            let generator_control_arity = usize::from(m.is_generator && !has_rest);
            if has_rest {
                cc.rest_fixed_arities
                    .insert(user_params.len().saturating_sub(1) as u8);
            }
            let arity = if is_static_init {
                (user_params.len() + generator_control_arity) as u8
            } else if is_static {
                if cc.profile.name == "php" {
                    (user_params.len() + 1 + generator_control_arity) as u8
                } else {
                    (user_params.len() + generator_control_arity) as u8
                }
            } else if uses_js_this {
                (user_params.len() + generator_control_arity) as u8
            } else {
                (user_params.len() + 1 + generator_control_arity) as u8
            };

            let ci = cc.chunks.len();
            let mut chunk = common::functions::create_function_chunk(mname, arity);
            chunk.is_async = m.is_async;
            chunk.is_generator = m.is_generator;
            if m.is_generator && !m.is_async {
                let cname_canon = cc.canon(mname);
                cc.generator_functions.insert(cname_canon);
            }
            if let Some(&n) = cc.multi_return_functions.get(&bound_name) {
                chunk.result_arity = n;
            }
            cc.chunks.push(chunk);
            cc.scopes.push(Scope::new_function());
            cc.static_local_bindings.push(HashMap::new());
            let saved = cc.current;
            cc.current = ci;
            let saved_fn = cc.current_func_name.take();
            cc.current_func_name = Some(bound_name.clone());

            if !uses_js_this && !is_static_init && (!is_static || cc.profile.name == "php") {
                cc.define_local(&self_kw);
            }
            for p in &user_params {
                cc.define_local_typed(&p.name, p.type_hint.clone());
                let normalized_type_hint =
                    p.type_hint.as_deref().map(Compiler::normalize_type_hint);
                if normalized_type_hint
                    .as_deref()
                    .is_some_and(|type_hint| type_hint.ends_with("()"))
                    || normalized_type_hint
                        .as_deref()
                        .is_some_and(|type_hint| type_hint.ends_with("()"))
                {
                    cc.record_array_binding(
                        &p.name,
                        ArrayBindingMetadata {
                            is_fixed: false,
                            type_hint: p.type_hint.clone(),
                            pascal_bounds: p.type_hint.as_deref().and_then(|type_hint| {
                                cc.pascal_array_type_hint_metadata(type_hint)
                            }),
                        },
                    );
                }
            }
            if !uses_js_this && !is_static_init && (!is_static || cc.profile.name == "php") {
                if class.explicit_self_param {
                    if let Some(self_param) = m.params.first() {
                        if self_param.name != self_kw {
                            let self_slot = cc.scope().resolve(&self_kw).unwrap();
                            let alias_slot = cc
                                .define_local_typed(&self_param.name, self_param.type_hint.clone());
                            cc.emit_u16(Op::LOCAL_GET, self_slot);
                            cc.emit_u16(Op::LOCAL_SET, alias_slot);
                            cc.emit(Op::DROP);
                        }
                    }
                }
            }
            let generator_control_slot = m
                .is_generator
                .then(|| cc.define_local("__generator_entry_control"));
            let ref_out_slots: Vec<u16> = user_params
                .iter()
                .filter(|param| matches!(param.pass_by, PassBy::Ref | PassBy::Out))
                .filter_map(|param| cc.scope().resolve(&param.name))
                .collect();
            let saved_rs = cc.current_result_slot.take();
            let saved_ref_out = cc.current_ref_out_params.take();
            let saved_member_static = cc.current_member_is_static;
            cc.current_result_slot = None;
            cc.current_ref_out_params = (!ref_out_slots.is_empty()).then_some(ref_out_slots);
            cc.current_member_is_static = is_static;

            if let Some(control_slot) = generator_control_slot {
                cc.emit_generator_entry_control(control_slot)?;
            }

            // Default parameters (C# `string greeting = "Hello"`): if
            // the slot is null/undefined when the method runs, install
            // the default. JS profile uses `REF_IS_UNDEFINED` (only
            // explicit `undefined` triggers); other languages use
            // `REF_IS_NULL` which matches either tag.
            for p in &user_params {
                if let Some(ref default) = p.default {
                    let slot = cc.scope().resolve(&p.name).unwrap();
                    cc.emit_u16(Op::LOCAL_GET, slot);
                    if uses_js_this {
                        cc.emit(Op::REF_IS_UNDEFINED);
                    } else {
                        cc.emit(Op::REF_IS_NULL);
                    }
                    let branch_line = cc.line;
                    crate::emitter::ops::emit_dyn_to_bool(cc.chunk(), branch_line);
                    cc.chunks[cc.current].emit_if(branch_line);
                    cc.compile_expr(default)?;
                    cc.emit_u16(Op::LOCAL_SET, slot);
                    cc.emit(Op::DROP);
                    cc.chunks[cc.current].emit_end(branch_line);
                }
            }

            if is_ctor {
                for s in &m.body {
                    cc.compile_stmt(s)?;
                }
                if let Some(slot) = cc
                    .scope()
                    .resolve(&self_kw)
                    .or_else(|| cc.scope().resolve_ci(&self_kw))
                {
                    cc.emit_u16(Op::LOCAL_GET, slot);
                    cc.emit_return_through_finally(1)?;
                }
            } else if m.return_type.is_some() && result_style == ReturnStyle::ResultSlot {
                let slot_name = cc.profile.result_slot_name.clone();
                let rs = cc.define_local(&slot_name);
                let returns_self_type = m
                    .return_type
                    .as_deref()
                    .is_some_and(|rt| rt.eq_ignore_ascii_case(&class.name));
                if returns_self_type && body_has_result_member_assign(&m.body) {
                    cc.emit_var_get(&class.name);
                    cc.emit_u8(Op::CALL_REF, 0);
                } else {
                    cc.emit(Op::NULL);
                }
                cc.emit_u16(Op::LOCAL_SET, rs);
                cc.emit(Op::DROP);
                cc.current_result_slot = Some(rs);
                for s in &m.body {
                    cc.compile_stmt(s)?;
                }
                cc.emit_u16(Op::LOCAL_GET, rs);
                cc.emit_return_through_finally(1)?;
            } else {
                for s in &m.body {
                    cc.compile_stmt(s)?;
                }
                if cc.current_ref_out_params.is_some() {
                    cc.emit(Op::NULL);
                    cc.emit_return_through_finally(1)?;
                } else {
                    let line = cc.line;
                    common::functions::emit_function_epilogue(&mut cc.chunks[ci], line);
                }
            }

            cc.current_func_name = saved_fn;
            cc.current_result_slot = saved_rs;
            cc.current_ref_out_params = saved_ref_out;
            cc.current_member_is_static = saved_member_static;

            let locals = cc.scope().next_slot.max(cc.chunks[ci].local_count);
            cc.chunks[ci].local_count = locals;
            let method_scope_idx = cc.scopes.len() - 1;
            let capture_names: Vec<String> = cc.scopes[method_scope_idx]
                .upvalues
                .iter()
                .enumerate()
                .filter_map(|(index, _)| {
                    cc.captured_name_for_upvalue(method_scope_idx, index as u8)
                })
                .collect();
            cc.scopes.pop();
            cc.static_local_bindings.pop();
            cc.current = saved;
            method_capture_name_map.insert(ci, capture_names);
            if let Some(pending) = cc.pending_classes.get_mut(name) {
                let overloads = if is_static {
                    &mut pending.static_method_overloads
                } else {
                    &mut pending.instance_method_overloads
                };
                overloads
                    .entry(bound_name.clone())
                    .or_default()
                    .push(PendingMethodOverload {
                        param_types,
                        chunk_idx: ci,
                        return_type: m.return_type.clone(),
                        signature: CallSignature::from_params(
                            &user_params
                                .iter()
                                .map(|param| (*param).clone())
                                .collect::<Vec<_>>(),
                        ),
                    });
            }
            method_chunks.push((bound_name, ci, is_ctor, is_static));
            Ok(())
        };

        for m in &class.instance_methods {
            compile_normal_method(self, m, false)?;
        }
        for m in &class.static_methods {
            compile_normal_method(self, m, true)?;
        }

        // --- Events / Consts / NestedTypes (from raw_extra_members) ---
        for m in &class.raw_extra_members {
            match m {
                ClassMember::Const {
                    name: cname, value, ..
                } => {
                    self.compile_expr(value)?;
                    let global_name = self.canon(&format!("{}.{}", name, cname));
                    let idx = self.str_const(&global_name);
                    self.emit_u16(Op::GLOBAL_SET, idx);
                    self.emit(Op::DROP);
                    self.defined_globals.insert(global_name);
                }
                ClassMember::Event { .. } => { /* type-level only */ }
                ClassMember::NestedType(stmt) => {
                    self.compile_stmt(stmt)?;
                }
                _ => {}
            }
        }

        // --- Properties: getter → __get_<prop>, setter → __set_<prop> ---
        for p in &class.properties {
            // Auto-properties are handled as plain fields in pass-1.
            if p.auto_field.is_some() {
                continue;
            }
            let pname_canon = if let Some(private_name) =
                self.js_private_member_storage_name_for_class(&class.name, &p.source_name)
            {
                private_name
            } else if !p.canonical_name.is_empty() {
                p.canonical_name.clone()
            } else {
                self.canon(&p.source_name)
            };
            let prop_is_static = p.is_static;

            if let Some(getter) = &p.getter {
                let get_name = format!("__get_{}", pname_canon);
                let ci = self.chunks.len();
                let chunk = common::functions::create_function_chunk(
                    &get_name,
                    if prop_is_static { 0 } else { 1 },
                );
                self.chunks.push(chunk);
                self.scopes.push(Scope::new_function());
                let saved = self.current;
                self.current = ci;
                if !prop_is_static {
                    self.define_local(&self_kw);
                }

                if getter.body.is_empty() {
                    // Auto-property getter: return backing field
                    if let Some(slot) = self.scope().resolve(&self_kw) {
                        self.emit_u16(Op::LOCAL_GET, slot);
                        let backing = self.str_const(&format!("__{}", pname_canon));
                        self.emit_u16(Op::STRUCT_GET, backing);
                        self.emit(Op::RETURN);
                    }
                } else {
                    let slot_name = self.profile.result_slot_name.clone();
                    let rs = self.define_local(&slot_name);
                    self.emit(Op::NULL);
                    self.emit_u16(Op::LOCAL_SET, rs);
                    self.emit(Op::DROP);
                    let saved_fn = self.current_func_name.take();
                    let saved_rs = self.current_result_slot.take();
                    self.current_func_name = Some(p.source_name.clone());
                    self.current_result_slot = Some(rs);
                    for s in &getter.body {
                        self.compile_stmt(s)?;
                    }
                    self.current_func_name = saved_fn;
                    self.current_result_slot = saved_rs;
                    self.emit_u16(Op::LOCAL_GET, rs);
                    self.emit(Op::RETURN);
                }

                let locals = self.scope().next_slot.max(self.chunks[ci].local_count);
                self.chunks[ci].local_count = locals;
                self.scopes.pop();
                self.current = saved;
                method_chunks.push((get_name, ci, false, prop_is_static));
            }

            if let Some(setter) = &p.setter {
                let set_name = format!("__set_{}", pname_canon);
                let ci = self.chunks.len();
                let chunk = common::functions::create_function_chunk(
                    &set_name,
                    if prop_is_static { 1 } else { 2 },
                );
                self.chunks.push(chunk);
                self.scopes.push(Scope::new_function());
                let saved = self.current;
                self.current = ci;
                if !prop_is_static {
                    self.define_local(&self_kw);
                }
                let value_param_name = setter
                    .params
                    .first()
                    .map(|p| p.name.clone())
                    .unwrap_or_else(|| "value".to_string());
                self.define_local(&value_param_name);

                if setter.body.is_empty() {
                    // Auto-property setter: set backing field
                    if let Some(self_slot) = self.scope().resolve(&self_kw) {
                        self.emit_u16(Op::LOCAL_GET, self_slot);
                        if let Some(val_slot) = self.scope().resolve(&value_param_name) {
                            self.emit_u16(Op::LOCAL_GET, val_slot);
                        }
                        let backing = self.str_const(&format!("__{}", pname_canon));
                        self.emit_u16(Op::STRUCT_SET, backing);
                        self.emit(Op::DROP);
                    }
                } else {
                    for s in &setter.body {
                        self.compile_stmt(s)?;
                    }
                }

                let line = self.line;
                common::functions::emit_function_epilogue(&mut self.chunks[ci], line);
                let locals = self.scope().next_slot.max(self.chunks[ci].local_count);
                self.chunks[ci].local_count = locals;
                self.scopes.pop();
                self.current = saved;
                method_chunks.push((set_name, ci, false, prop_is_static));
            }
        }

        self.current_class = saved_class;
        self.current_class_implicit_self = saved_implicit;

        const IMPLICIT_CTOR_FORWARD_ARGS: u8 = 16;
        let instance_methods: Vec<&(String, usize, bool, bool)> = method_chunks
            .iter()
            .filter(|(_, _, ic, is_static)| !*ic && !*is_static)
            .collect();
        let static_methods: Vec<&(String, usize, bool, bool)> = method_chunks
            .iter()
            .filter(|(_, _, ic, is_static)| !*ic && *is_static)
            .collect();
        let instance_method_names: Vec<String> = instance_methods
            .iter()
            .map(|(n, _, _, _)| n.clone())
            .collect();
        let method_rest_fixed_counts: HashMap<usize, u8> = self
            .pending_classes
            .values()
            .flat_map(|pc| {
                pc.instance_method_overloads
                    .values()
                    .chain(pc.static_method_overloads.values())
            })
            .flat_map(|overloads| overloads.iter())
            .filter(|overload| overload.signature.has_rest)
            .map(|overload| {
                (
                    overload.chunk_idx,
                    overload.signature.param_names.len().saturating_sub(1) as u8,
                )
            })
            .collect();
        let method_rest_fixed_count =
            |chunk_idx: usize| method_rest_fixed_counts.get(&chunk_idx).copied();

        let ctor_variants: Vec<Option<&NormalConstructor>> = if !class.constructors.is_empty() {
            class.constructors.iter().map(Some).collect()
        } else if let Some(ctor) = class.constructor.as_ref() {
            vec![Some(ctor)]
        } else {
            vec![None]
        };
        let ctor_global_prefix = self.canon(name);
        let should_stamp_form_identity = self.class_requires_form_identity_stamp(parent);
        for ctor_variant in &ctor_variants {
            let explicit_arity = ctor_variant
                .map(|ctor| {
                    let skip = if class.explicit_self_param { 1 } else { 0 };
                    ctor.params.len().saturating_sub(skip)
                })
                .unwrap_or_else(|| {
                    if parent.is_some() {
                        IMPLICIT_CTOR_FORWARD_ARGS as usize
                    } else {
                        0
                    }
                });
            self.defined_globals
                .insert(format!("{}$arity{}", ctor_global_prefix, explicit_arity));
        }

        let emit_helper_ref =
            |cc: &mut Compiler, helper_idx: usize, helper_upvalues: &[UpvalueDesc], line: u32| {
                common::functions::emit_ref_func(
                    &mut cc.chunks[cc.current],
                    helper_idx,
                    helper_upvalues.len() as u8,
                    line,
                );
                for uv in helper_upvalues {
                    cc.chunks[cc.current].emit(if uv.is_local { 1 } else { 0 }, line);
                    cc.chunks[cc.current].emit(uv.index, line);
                }
            };

        let mut ctor_helpers: Vec<(usize, usize, usize, Vec<UpvalueDesc>)> = Vec::new();
        for (ctor_index, ctor_variant) in ctor_variants.iter().enumerate() {
            let helper_name = format!("__{}_ctor_{}", name, ctor_index);
            let ctor_base_args_from_nc: Option<Vec<Expression>> = ctor_variant.and_then(|c| {
                if let BaseCall::Explicit(args) = &c.base_call {
                    Some(args.iter().map(|a| a.value.clone()).collect())
                } else {
                    None
                }
            });
            let ctor_auto_base = ctor_variant
                .map(|c| matches!(c.base_call, BaseCall::Auto))
                .unwrap_or(false);
            let ctor_this_args: Option<Vec<Expression>> = ctor_variant.and_then(|c| {
                if let BaseCall::This(args) = &c.base_call {
                    Some(args.iter().map(|a| a.value.clone()).collect())
                } else {
                    None
                }
            });
            let ctor_body: Option<(&Vec<Statement>, &Vec<Param>, Option<&Vec<Expression>>)> =
                ctor_variant.map(|c| (&c.body, &c.params, ctor_base_args_from_nc.as_ref()));
            if let Some((_, params, _)) = ctor_body {
                let skip = if class.explicit_self_param { 1 } else { 0 };
                let ctor_params: Vec<Param> = params.iter().skip(skip).cloned().collect();
                self.constructor_signatures
                    .entry(self.canon(name))
                    .or_default()
                    .push(CallSignature::from_params(&ctor_params));
            }
            let user_params: Vec<String> = ctor_body
                .map(|(_, params, _)| {
                    if class.explicit_self_param {
                        params.iter().skip(1).map(|p| p.name.clone()).collect()
                    } else {
                        params.iter().map(|p| p.name.clone()).collect()
                    }
                })
                .unwrap_or_default();
            let ctor_min_arity = ctor_body
                .map(|(_, params, _)| {
                    let skip = if class.explicit_self_param { 1 } else { 0 };
                    params
                        .iter()
                        .skip(skip)
                        .take_while(|p| p.default.is_none() && !p.is_rest)
                        .count()
                })
                .unwrap_or(0);
            let synthesized_forward_args = ctor_body.is_none() && parent.is_some();
            let user_arity = if synthesized_forward_args {
                IMPLICIT_CTOR_FORWARD_ARGS
            } else {
                user_params.len() as u8
            };

            let helper_idx = self.chunks.len();
            self.chunks.push(common::functions::create_function_chunk(
                &helper_name,
                user_arity,
            ));
            self.scopes.push(Scope::new_function());
            let saved_cur = self.current;
            let saved_class2 = self.current_class.take();
            let saved_implicit2 = self.current_class_implicit_self;
            self.current = helper_idx;
            self.current_class = Some(name.to_string());
            self.current_class_implicit_self = class.implicit_self_fields;

            let ctor_param_defaults: Vec<Option<Expression>> = ctor_body
                .map(|(_, params, _)| {
                    let skip = if class.explicit_self_param { 1 } else { 0 };
                    params
                        .iter()
                        .skip(skip)
                        .map(|p| p.default.clone())
                        .collect()
                })
                .unwrap_or_default();
            for p in &user_params {
                self.define_local(p);
            }
            for (i, p) in user_params.iter().enumerate() {
                if let Some(Some(default)) = ctor_param_defaults.get(i) {
                    let slot = self.scope().resolve(p).unwrap();
                    self.emit_u16(Op::LOCAL_GET, slot);
                    if self.is_js_profile() {
                        self.emit(Op::REF_IS_UNDEFINED);
                    } else {
                        self.emit(Op::REF_IS_NULL);
                    }
                    let branch_line = self.line;
                    crate::emitter::ops::emit_dyn_to_bool(self.chunk(), branch_line);
                    self.chunks[self.current].emit_if(branch_line);
                    self.compile_expr(default)?;
                    self.emit_u16(Op::LOCAL_SET, slot);
                    self.emit(Op::DROP);
                    self.chunks[self.current].emit_end(branch_line);
                }
            }
            if synthesized_forward_args {
                for i in 0..IMPLICIT_CTOR_FORWARD_ARGS {
                    self.define_local(&format!("__implicit_arg_{}", i));
                }
            }
            self.define_local(&self_kw);
            let this_slot = user_arity as u16;
            if self.is_js_profile() {
                let js_this = self.str_const("__js_this");
                self.emit_u16(Op::GLOBAL_GET, js_this);
                self.emit_u16(Op::LOCAL_SET, this_slot);
                self.emit(Op::DROP);
            }

            let line = self.line;
            if let Some(this_args) = ctor_this_args {
                let ctor_global = format!("{}$arity{}", ctor_global_prefix, this_args.len());
                self.emit_var_get(&ctor_global);
                for expr in &this_args {
                    self.compile_expr(expr)?;
                }
                self.emit_u8(Op::CALL_REF, this_args.len() as u8);
                self.emit_u16(Op::LOCAL_SET, this_slot);
                self.emit(Op::DROP);
                if let Some((body, _, _)) = ctor_body {
                    for stmt in body {
                        self.compile_stmt(stmt)?;
                    }
                }
                for (mname, mci, _, _) in &instance_methods {
                    if mname.starts_with("__get_") {
                        let prop = mname.strip_prefix("__get_").unwrap_or(mname);
                        common::classes::emit_bind_getter(
                            self.chunk(),
                            this_slot,
                            prop,
                            *mci,
                            line,
                        );
                    } else if mname.starts_with("__set_") {
                        let prop = mname.strip_prefix("__set_").unwrap_or(mname);
                        common::classes::emit_bind_setter(
                            self.chunk(),
                            this_slot,
                            prop,
                            *mci,
                            line,
                        );
                    } else {
                        let capture_names = method_capture_name_map
                            .get(mci)
                            .map(Vec::as_slice)
                            .unwrap_or(&[]);
                        self.emit_bind_instance_method_with_aliases(
                            this_slot,
                            mname,
                            *mci,
                            capture_names,
                            method_rest_fixed_count(*mci),
                            !self.is_js_profile(),
                        )?;
                    }
                }
                common::classes::emit_constructor_return(self.chunk(), this_slot, line);
            } else {
                let is_child = parent.is_some();
                let parent_ctor_is_bound = if let Some(parent_name) = parent {
                    let pname = self.canon(parent_name);
                    let has_local = self
                        .scope()
                        .resolve(parent_name)
                        .or_else(|| {
                            if self.case_sensitive {
                                None
                            } else {
                                self.scope().resolve_ci(parent_name)
                            }
                        })
                        .is_some();
                    let has_upvalue = self.scopes.len() > 1
                        && self
                            .resolve_upvalue(self.scopes.len() - 1, parent_name)
                            .is_some();
                    let has_static_local = self.static_local_binding(parent_name).is_some();
                    has_local
                        || has_upvalue
                        || has_static_local
                        || self.defined_globals.contains(&pname)
                        || self.defined_classes.contains(&pname)
                } else {
                    false
                };
                if is_child {
                    self.emit(Op::NULL);
                    self.emit_u16(Op::LOCAL_SET, this_slot);
                    self.emit(Op::DROP);

                    let has_explicit_base =
                        ctor_body.as_ref().map_or(false, |(_, _, ba)| ba.is_some());
                    let auto_base_needed = !has_explicit_base
                        && ctor_body.is_some()
                        && ctor_auto_base
                        && parent.is_some()
                        && {
                            let stmts = ctor_body
                                .as_ref()
                                .map(|(b, _, _)| b.as_slice())
                                .unwrap_or(&[]);
                            !body_has_super_call(stmts)
                        };

                    if let Some((_, _, base_args)) = &ctor_body {
                        if let Some(bargs) = base_args {
                            if let Some(parent_name) = parent {
                                if parent_ctor_is_bound {
                                    if self.is_js_profile() {
                                        self.emit_var_get(name);
                                        self.set_js_new_target_from_stack();
                                    }
                                    self.emit_var_get(parent_name);
                                    for a in *bargs {
                                        self.compile_expr(a)?;
                                    }
                                    self.emit_u8(Op::CALL_REF, bargs.len() as u8);
                                    self.emit_u16(Op::LOCAL_SET, this_slot);
                                    self.emit(Op::DROP);
                                } else {
                                    let canon_name = self.canon(name);
                                    common::classes::emit_new_typed_object(
                                        self.chunk(),
                                        this_slot,
                                        &canon_name,
                                        line,
                                    );
                                }
                            }
                        } else if auto_base_needed {
                            if let Some(parent_name) = parent {
                                if parent_ctor_is_bound {
                                    if self.is_js_profile() {
                                        self.emit_var_get(name);
                                        self.set_js_new_target_from_stack();
                                    }
                                    self.emit_var_get(parent_name);
                                    self.emit_u8(Op::CALL_REF, 0);
                                    self.emit_u16(Op::LOCAL_SET, this_slot);
                                    self.emit(Op::DROP);
                                } else {
                                    let canon_name = self.canon(name);
                                    common::classes::emit_new_typed_object(
                                        self.chunk(),
                                        this_slot,
                                        &canon_name,
                                        line,
                                    );
                                }
                            }
                        }
                    } else if let Some(parent_name) = parent {
                        if parent_ctor_is_bound {
                            if self.is_js_profile() {
                                self.emit_var_get(name);
                                self.set_js_new_target_from_stack();
                            }
                            self.emit_var_get(parent_name);
                            if synthesized_forward_args {
                                let parent_ctor_slot =
                                    self.define_local(&format!("__{}_parent_ctor", helper_name));
                                self.emit_u16(Op::LOCAL_SET, parent_ctor_slot);
                                self.emit(Op::DROP);
                                let parent_called_slot =
                                    self.define_local(&format!("__{}_parent_called", helper_name));
                                self.emit(Op::I32_CONST_0);
                                self.emit_u16(Op::LOCAL_SET, parent_called_slot);
                                self.emit(Op::DROP);
                                for count in (1..=IMPLICIT_CTOR_FORWARD_ARGS).rev() {
                                    self.emit_u16(Op::LOCAL_GET, parent_called_slot);
                                    self.emit(Op::I32_EQZ);
                                    self.chunks[self.current].emit_if(line);
                                    self.emit_u16(Op::LOCAL_GET, (count - 1) as u16);
                                    self.emit(Op::REF_IS_NULL);
                                    let branch_line = self.line;
                                    crate::emitter::ops::emit_dyn_to_bool(
                                        self.chunk(),
                                        branch_line,
                                    );
                                    self.emit(Op::I32_EQZ);
                                    self.chunks[self.current].emit_if(line);
                                    self.emit_u16(Op::LOCAL_GET, parent_ctor_slot);
                                    for arg_index in 0..count {
                                        self.emit_u16(Op::LOCAL_GET, arg_index as u16);
                                    }
                                    self.emit_u8(Op::CALL_REF, count);
                                    self.emit_u16(Op::LOCAL_SET, this_slot);
                                    self.emit(Op::DROP);
                                    self.emit(Op::I32_CONST_1);
                                    self.emit_u16(Op::LOCAL_SET, parent_called_slot);
                                    self.emit(Op::DROP);
                                    self.chunks[self.current].emit_end(line);
                                    self.chunks[self.current].emit_end(line);
                                }
                                self.emit_u16(Op::LOCAL_GET, parent_called_slot);
                                self.emit(Op::I32_EQZ);
                                self.chunks[self.current].emit_if(line);
                                self.emit_u16(Op::LOCAL_GET, parent_ctor_slot);
                                self.emit_u8(Op::CALL_REF, 0);
                                self.emit_u16(Op::LOCAL_SET, this_slot);
                                self.emit(Op::DROP);
                                self.chunks[self.current].emit_end(line);
                            } else {
                                for i in 0..user_arity {
                                    self.emit_u16(Op::LOCAL_GET, i as u16);
                                }
                                self.emit_u8(Op::CALL_REF, user_arity);
                                self.emit_u16(Op::LOCAL_SET, this_slot);
                                self.emit(Op::DROP);
                            }
                        } else {
                            let canon_name = self.canon(name);
                            common::classes::emit_new_typed_object(
                                self.chunk(),
                                this_slot,
                                &canon_name,
                                line,
                            );
                        }
                    }

                    if has_explicit_base || auto_base_needed || ctor_body.is_none() {
                        self.emit_u16(Op::LOCAL_GET, this_slot);
                        self.emit_const(Value::String(Arc::from(name)));
                        let type_key = self.str_const("__type");
                        self.emit_u16(Op::STRUCT_SET, type_key);
                        self.emit(Op::DROP);

                        if self.is_js_profile() {
                            let class_global = self.str_const(name);
                            let prototype_key = self.str_const("prototype");
                            let proto_link_key = self.str_const("__proto__");
                            let proto_local =
                                self.define_local(&format!("__{}_link_proto", helper_name));
                            self.emit_u16(Op::GLOBAL_GET, class_global);
                            self.emit_u16(Op::STRUCT_GET, prototype_key);
                            self.emit_u16(Op::LOCAL_SET, proto_local);
                            self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_GET, proto_local);
                            self.emit(Op::REF_IS_NULL);
                            let branch_line = self.line;
                            crate::emitter::ops::emit_dyn_to_bool(self.chunk(), branch_line);
                            self.emit(Op::I32_EQZ);
                            self.chunks[self.current].emit_if(line);
                            self.emit_u16(Op::LOCAL_GET, this_slot);
                            self.emit_u16(Op::LOCAL_GET, proto_local);
                            self.emit_u16(Op::STRUCT_SET, proto_link_key);
                            self.emit(Op::DROP);
                            self.chunks[self.current].emit_end(line);
                        }

                        for (fname, type_hint, init, array_bounds) in &field_inits {
                            self.emit_class_field_initializer(
                                this_slot,
                                fname,
                                type_hint.as_deref(),
                                init.as_ref(),
                                array_bounds.as_deref(),
                                class.is_value_type,
                                line,
                            )?;
                        }

                        if let Some(parent_name) = parent {
                            let pname = self.canon(parent_name);
                            for method_name in &instance_method_names {
                                common::classes::emit_save_base_method(
                                    self.chunk(),
                                    this_slot,
                                    method_name,
                                    line,
                                );
                            }
                            self.emit_store_super_ref(this_slot, &pname);
                        }

                        for (mname, mci, _, _) in &instance_methods {
                            if mname.starts_with("__get_") {
                                let prop = mname.strip_prefix("__get_").unwrap_or(mname);
                                common::classes::emit_bind_getter(
                                    self.chunk(),
                                    this_slot,
                                    prop,
                                    *mci,
                                    line,
                                );
                            } else if mname.starts_with("__set_") {
                                let prop = mname.strip_prefix("__set_").unwrap_or(mname);
                                common::classes::emit_bind_setter(
                                    self.chunk(),
                                    this_slot,
                                    prop,
                                    *mci,
                                    line,
                                );
                            } else {
                                let capture_names = method_capture_name_map
                                    .get(mci)
                                    .map(Vec::as_slice)
                                    .unwrap_or(&[]);
                                self.emit_bind_instance_method_with_aliases(
                                    this_slot,
                                    mname,
                                    *mci,
                                    capture_names,
                                    method_rest_fixed_count(*mci),
                                    !self.is_js_profile(),
                                )?;
                            }
                        }

                        let ctor_stmts: &[Statement] = ctor_body
                            .as_ref()
                            .map(|(b, _, _)| b.as_slice())
                            .unwrap_or(&[]);
                        if should_stamp_form_identity && !body_has_identity_stamp(ctor_stmts) {
                            self.emit_form_identity_stamp(this_slot, name, line);
                        }
                        for aim in &class.auto_init_methods {
                            let has_method = instance_methods
                                .iter()
                                .any(|(n, _, _, _)| n.eq_ignore_ascii_case(aim));
                            if has_method && !body_calls_method(ctor_stmts, aim) {
                                common::classes::emit_auto_init_call(
                                    self.chunk(),
                                    this_slot,
                                    aim,
                                    line,
                                );
                            }
                        }

                        if let Some((body, _, _)) = ctor_body {
                            for stmt in body {
                                self.compile_stmt(stmt)?;
                            }
                        }
                    } else {
                        let body_stmts: &[Statement] = ctor_body
                            .as_ref()
                            .map(|(b, _, _)| b.as_slice())
                            .unwrap_or(&[]);
                        let is_super_call = |stmt: &Statement| {
                            if let StmtKind::Expr(expr) = &stmt.kind {
                                matches!(&expr.kind, ExprKind::SuperCall { .. })
                                    || matches!(&expr.kind, ExprKind::Call { callee, .. } if matches!(callee.kind, ExprKind::Super))
                            } else {
                                false
                            }
                        };
                        let super_idx = body_stmts.iter().position(is_super_call);
                        let preamble_end = match super_idx {
                            Some(index) => {
                                let mut end = index + 1;
                                while end < body_stmts.len() && is_identity_stamp(&body_stmts[end])
                                {
                                    end += 1;
                                }
                                end
                            }
                            None => 0,
                        };
                        for stmt in &body_stmts[..preamble_end] {
                            self.compile_stmt(stmt)?;
                        }
                        self.emit_u16(Op::LOCAL_GET, this_slot);
                        self.emit_const(Value::String(Arc::from(name)));
                        let type_key2 = self.str_const("__type");
                        self.emit_u16(Op::STRUCT_SET, type_key2);
                        self.emit(Op::DROP);
                        let tid_key = self.str_const(&format!("__tid_{}", self.canon(name)));
                        self.emit_u16(Op::LOCAL_GET, this_slot);
                        self.emit_u16(Op::GLOBAL_GET, tid_key);
                        self.emit(Op::SET_TYPE_ID);
                        self.emit(Op::DROP);
                        if let Some(parent_name) = parent {
                            let pname = self.canon(parent_name);
                            for method_name in &instance_method_names {
                                common::classes::emit_save_base_method(
                                    self.chunk(),
                                    this_slot,
                                    method_name,
                                    line,
                                );
                            }
                            self.emit_store_super_ref(this_slot, &pname);
                        }
                        for (fname, type_hint, init, array_bounds) in &field_inits {
                            self.emit_class_field_initializer(
                                this_slot,
                                fname,
                                type_hint.as_deref(),
                                init.as_ref(),
                                array_bounds.as_deref(),
                                class.is_value_type,
                                line,
                            )?;
                        }
                        for (mname, mci, _, _) in &instance_methods {
                            if mname.starts_with("__get_") {
                                let prop = mname.strip_prefix("__get_").unwrap_or(mname);
                                common::classes::emit_bind_getter(
                                    self.chunk(),
                                    this_slot,
                                    prop,
                                    *mci,
                                    line,
                                );
                            } else if mname.starts_with("__set_") {
                                let prop = mname.strip_prefix("__set_").unwrap_or(mname);
                                common::classes::emit_bind_setter(
                                    self.chunk(),
                                    this_slot,
                                    prop,
                                    *mci,
                                    line,
                                );
                            } else {
                                let capture_names = method_capture_name_map
                                    .get(mci)
                                    .map(Vec::as_slice)
                                    .unwrap_or(&[]);
                                self.emit_bind_instance_method_with_aliases(
                                    this_slot,
                                    mname,
                                    *mci,
                                    capture_names,
                                    method_rest_fixed_count(*mci),
                                    !self.is_js_profile(),
                                )?;
                            }
                        }
                        let user_body = &body_stmts[preamble_end..];
                        if should_stamp_form_identity && !body_has_identity_stamp(body_stmts) {
                            self.emit_form_identity_stamp(this_slot, name, line);
                        }
                        for aim in &class.auto_init_methods {
                            let has_method = instance_methods
                                .iter()
                                .any(|(n, _, _, _)| n.eq_ignore_ascii_case(aim));
                            if has_method && !body_calls_method(user_body, aim) {
                                common::classes::emit_auto_init_call(
                                    self.chunk(),
                                    this_slot,
                                    aim,
                                    line,
                                );
                            }
                        }
                        for stmt in user_body {
                            self.compile_stmt(stmt)?;
                        }
                    }
                } else {
                    let canon_name = self.canon(name);
                    common::classes::emit_new_typed_object(
                        self.chunk(),
                        this_slot,
                        &canon_name,
                        line,
                    );
                    for (fname, type_hint, init, array_bounds) in &field_inits {
                        self.emit_class_field_initializer(
                            this_slot,
                            fname,
                            type_hint.as_deref(),
                            init.as_ref(),
                            array_bounds.as_deref(),
                            class.is_value_type,
                            line,
                        )?;
                    }
                    for (mname, mci, _, _) in &instance_methods {
                        if mname.starts_with("__get_") {
                            let prop = mname.strip_prefix("__get_").unwrap_or(mname);
                            common::classes::emit_bind_getter(
                                self.chunk(),
                                this_slot,
                                prop,
                                *mci,
                                line,
                            );
                        } else if mname.starts_with("__set_") {
                            let prop = mname.strip_prefix("__set_").unwrap_or(mname);
                            common::classes::emit_bind_setter(
                                self.chunk(),
                                this_slot,
                                prop,
                                *mci,
                                line,
                            );
                        } else {
                            let capture_names = method_capture_name_map
                                .get(mci)
                                .map(Vec::as_slice)
                                .unwrap_or(&[]);
                            self.emit_bind_instance_method_with_aliases(
                                this_slot,
                                mname,
                                *mci,
                                capture_names,
                                method_rest_fixed_count(*mci),
                                !self.is_js_profile(),
                            )?;
                        }
                    }
                    if self.is_js_profile() {
                        let class_global = self.str_const(name);
                        let prototype_key = self.str_const("prototype");
                        let proto_link_key = self.str_const("__proto__");
                        let proto_local =
                            self.define_local(&format!("__{}_link_proto_base", helper_name));
                        self.emit_u16(Op::GLOBAL_GET, class_global);
                        self.emit_u16(Op::STRUCT_GET, prototype_key);
                        self.emit_u16(Op::LOCAL_SET, proto_local);
                        self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, proto_local);
                        self.emit(Op::REF_IS_NULL);
                        let branch_line = self.line;
                        crate::emitter::ops::emit_dyn_to_bool(self.chunk(), branch_line);
                        self.emit(Op::I32_EQZ);
                        self.chunks[self.current].emit_if(line);
                        self.emit_u16(Op::LOCAL_GET, this_slot);
                        self.emit_u16(Op::LOCAL_GET, proto_local);
                        self.emit_u16(Op::STRUCT_SET, proto_link_key);
                        self.emit(Op::DROP);
                        self.chunks[self.current].emit_end(line);
                    }
                    let ctor_stmts: &[Statement] = ctor_body
                        .as_ref()
                        .map(|(b, _, _)| b.as_slice())
                        .unwrap_or(&[]);
                    for aim in &class.auto_init_methods {
                        let has_method = instance_methods
                            .iter()
                            .any(|(n, _, _, _)| n.eq_ignore_ascii_case(aim));
                        if has_method && !body_calls_method(ctor_stmts, aim) {
                            common::classes::emit_auto_init_call(
                                self.chunk(),
                                this_slot,
                                aim,
                                line,
                            );
                        }
                    }
                    if let Some((body, _, _)) = ctor_body {
                        for stmt in body {
                            self.compile_stmt(stmt)?;
                        }
                    }
                }

                common::classes::emit_instanceof_chain(
                    &mut self.chunks,
                    self.current,
                    this_slot,
                    name,
                    line,
                );
                for interface_name in self.reflection_interfaces(name) {
                    common::classes::emit_instanceof_chain(
                        &mut self.chunks,
                        self.current,
                        this_slot,
                        &interface_name,
                        line,
                    );
                }
                common::classes::emit_constructor_return(self.chunk(), this_slot, line);
            }

            let locals = self
                .scope()
                .next_slot
                .max(self.chunks[helper_idx].local_count);
            self.chunks[helper_idx].local_count = locals;
            let helper_upvalues = self.scope().upvalues.clone();
            self.scopes.pop();
            self.current = saved_cur;
            self.current_class = saved_class2;
            self.current_class_implicit_self = saved_implicit2;
            ctor_helpers.push((
                user_arity as usize,
                ctor_min_arity,
                helper_idx,
                helper_upvalues,
            ));
        }

        let ctor_idx = self.chunks.len();
        let ctor_arity = ctor_helpers
            .iter()
            .map(|(arity, _, _, _)| *arity)
            .max()
            .unwrap_or(0) as u8;
        self.chunks
            .push(common::functions::create_function_chunk(name, ctor_arity));
        self.scopes.push(Scope::new_function());
        let saved_cur = self.current;
        self.current = ctor_idx;
        for i in 0..ctor_arity {
            self.define_local(&format!("__ctor_arg_{}", i));
        }
        let line = self.line;
        let js_ctor_relaxes_min_arity = self.is_js_profile();
        let helper_for_count = |count: usize| {
            ctor_helpers
                .iter()
                .filter(|(arity, min_arity, _, _)| {
                    count <= *arity && (js_ctor_relaxes_min_arity || count >= *min_arity)
                })
                .min_by_key(|(arity, _, _, _)| *arity)
        };
        for count in (1..=ctor_arity as usize).rev() {
            self.emit_u16(Op::LOCAL_GET, (count - 1) as u16);
            self.emit(Op::REF_IS_NULL);
            let branch_line = self.line;
            crate::emitter::ops::emit_dyn_to_bool(self.chunk(), branch_line);
            self.emit(Op::I32_EQZ);
            self.chunks[self.current].emit_if(line);
            if let Some((_, _, helper_idx, helper_upvalues)) = helper_for_count(count) {
                emit_helper_ref(self, *helper_idx, helper_upvalues, line);
                for arg_index in 0..count {
                    self.emit_u16(Op::LOCAL_GET, arg_index as u16);
                }
                self.emit_u8(Op::CALL_REF, count as u8);
                self.emit_return_through_finally(1)?;
            }
            self.chunks[self.current].emit_end(line);
        }
        if let Some((_, _, helper_idx, helper_upvalues)) = helper_for_count(0) {
            emit_helper_ref(self, *helper_idx, helper_upvalues, line);
            self.emit_u8(Op::CALL_REF, 0);
        } else {
            self.emit(Op::NULL);
        }
        self.emit_return_through_finally(1)?;
        let locals = self
            .scope()
            .next_slot
            .max(self.chunks[ctor_idx].local_count);
        self.chunks[ctor_idx].local_count = locals;
        let ctor_upvalues = self.scope().upvalues.clone();
        self.scopes.pop();
        self.current = saved_cur;

        let ctor_local = self.define_local(&format!("__{}_ctor", name));
        let uv_pairs: Vec<(bool, u8)> = ctor_upvalues
            .iter()
            .map(|uv| (uv.is_local, uv.index))
            .collect();
        let case_sensitive = self.is_js_profile();
        common::classes::emit_store_constructor_with_upvalues(
            self.chunk(),
            name,
            ctor_idx,
            ctor_local,
            &uv_pairs,
            case_sensitive,
            line,
        );
        for (arity, _, helper_idx, helper_upvalues) in &ctor_helpers {
            emit_helper_ref(self, *helper_idx, helper_upvalues, line);
            let helper_global = format!("{}$arity{}", ctor_global_prefix, arity);
            let helper_idx_const = self.str_const(&helper_global);
            self.emit_u16(Op::GLOBAL_SET, helper_idx_const);
        }

        if self.is_js_profile() {
            self.emit_common("object.new", 0, line);
            let proto_local = self.define_local(&format!("__{}_prototype", name));
            self.emit_u16(Op::LOCAL_SET, proto_local);
            self.emit(Op::DROP);

            if let Some(parent_name) = parent {
                let pname = self.canon(parent_name);
                self.emit_var_get(&pname);
                let parent_proto_key = self.str_const("prototype");
                self.emit_u16(Op::STRUCT_GET, parent_proto_key);
                let parent_proto_local = self.define_local(&format!("__{}_parent_prototype", name));
                self.emit_u16(Op::LOCAL_SET, parent_proto_local);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, parent_proto_local);
                self.emit(Op::REF_IS_NULL);
                let branch_line = self.line;
                crate::emitter::ops::emit_dyn_to_bool(self.chunk(), branch_line);
                self.emit(Op::I32_EQZ);
                self.chunks[self.current].emit_if(line);
                self.emit_u16(Op::LOCAL_GET, proto_local);
                self.emit_u16(Op::LOCAL_GET, parent_proto_local);
                let proto_link_key = self.str_const("__proto__");
                self.emit_u16(Op::STRUCT_SET, proto_link_key);
                self.emit(Op::DROP);
                self.chunks[self.current].emit_end(line);
            }

            self.emit_u16(Op::LOCAL_GET, proto_local);
            self.emit_u16(Op::LOCAL_GET, ctor_local);
            let ctor_key = self.str_const("constructor");
            self.emit_u16(Op::STRUCT_SET, ctor_key);
            self.emit(Op::DROP);

            self.emit_u16(Op::LOCAL_GET, ctor_local);
            self.emit_u16(Op::LOCAL_GET, proto_local);
            let proto_key = self.str_const("prototype");
            self.emit_u16(Op::STRUCT_SET, proto_key);
            self.emit(Op::DROP);

            self.emit_u16(Op::LOCAL_GET, ctor_local);
            self.emit_const(Value::String(Arc::from(name)));
            let name_key = self.str_const("name");
            self.emit_u16(Op::STRUCT_SET, name_key);
            self.emit(Op::DROP);
        }

        // Initialize static fields on the constructor object
        for (fname, type_hint, init, array_bounds) in &static_field_inits {
            self.emit_u16(Op::LOCAL_GET, ctor_local);
            if let Some(init_expr) = init {
                self.compile_expr(init_expr)?;
            } else if let Some(extent) = array_bounds
                .as_deref()
                .and_then(Self::array_bounds_extent_expr)
            {
                let init_expr = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("Array")),
                    args: vec![
                        Argument::positional(extent),
                        Argument::positional(Self::array_default_element_expr(
                            type_hint.as_deref(),
                        )),
                    ],
                    optional: false,
                });
                self.compile_expr(&init_expr)?;
            } else {
                self.emit(Op::NULL);
            }
            let fk = self.str_const(fname);
            self.emit_u16(Op::STRUCT_SET, fk);
            self.emit(Op::DROP);
        }

        for const_name in &static_const_names {
            self.emit_u16(Op::LOCAL_GET, ctor_local);
            let global_name = self.canon(&format!("{}.{}", name, const_name));
            let global_idx = self.str_const(&global_name);
            self.emit_u16(Op::GLOBAL_GET, global_idx);
            let field_idx = self.str_const(const_name);
            self.emit_u16(Op::STRUCT_SET, field_idx);
            self.emit(Op::DROP);
        }

        let own_static_member_names: Vec<String> = static_field_inits
            .iter()
            .map(|(fname, _, _, _)| fname.clone())
            .chain(static_const_names.iter().cloned())
            .collect();

        // Inherit static fields/constants onto the child class object so
        // PHP late static binding (`static::$x`, `static::NAME`) resolves
        // against the called class instead of stopping at the declaring class.
        if let Some(parent_name) = parent {
            let mut current_parent = Some(self.canon(parent_name));
            while let Some(ref pname) = current_parent {
                let parent_static_fields = self
                    .pending_classes
                    .get(pname.as_str())
                    .map(|pc| pc.static_fields.clone())
                    .unwrap_or_default();
                let next_parent = self
                    .pending_classes
                    .get(pname.as_str())
                    .and_then(|pc| pc.parent.clone());
                for field_name in &parent_static_fields {
                    if own_static_member_names
                        .iter()
                        .any(|name| name == field_name)
                    {
                        continue;
                    }
                    self.emit_u16(Op::LOCAL_GET, ctor_local);
                    let parent_idx = self.str_const(pname);
                    self.emit_u16(Op::GLOBAL_GET, parent_idx);
                    let field_idx = self.str_const(field_name);
                    self.emit_u16(Op::STRUCT_GET, field_idx);
                    self.emit_u16(Op::STRUCT_SET, field_idx);
                    self.emit(Op::DROP);
                }
                current_parent = next_parent;
            }
        }

        // Attach nested types onto the constructor object so `Outer.Inner`
        // resolves to the nested class constructor through the same shared
        // class-object path as static methods.
        let nested_types = self
            .pending_classes
            .get(name)
            .map(|pc| pc.nested_types.clone())
            .unwrap_or_default();
        for nested in nested_types {
            self.emit_u16(Op::LOCAL_GET, ctor_local);
            let nested_canon = self.canon(&nested);
            let nested_idx = self.str_const(&nested_canon);
            self.emit_u16(Op::GLOBAL_GET, nested_idx);
            let key = self.str_const(&nested_canon);
            self.emit_u16(Op::STRUCT_SET, key);
            self.emit(Op::DROP);
        }

        // Attach static methods to the constructor object
        let mut all_statics: Vec<(String, usize)> = Vec::new();
        let php_static_receiver = if self.profile.name == "php" {
            Some(ctor_local)
        } else {
            None
        };
        for (mname, mci, _, _) in &static_methods {
            common::classes::emit_attach_static_method(
                self.chunk(),
                ctor_local,
                mname,
                *mci,
                php_static_receiver,
                method_rest_fixed_count(*mci),
                line,
            );
            all_statics.push((mname.clone(), *mci));
        }

        if self.is_js_profile() {
            for method in &class.static_methods {
                if !method.source_name.starts_with('#') {
                    continue;
                }
                let bound_name =
                    self.js_member_storage_name_for_class(&class.name, &method.source_name);
                if let Some((_, chunk_idx, _, _)) =
                    method_chunks.iter().find(|(name, _, is_ctor, is_static)| {
                        !*is_ctor && *is_static && name == &bound_name
                    })
                {
                    common::classes::emit_attach_static_method(
                        self.chunk(),
                        ctor_local,
                        &bound_name,
                        *chunk_idx,
                        php_static_receiver,
                        method_rest_fixed_count(*chunk_idx),
                        line,
                    );
                }
            }
        }

        // Synthetic static constructor hook from language walkers.
        if let Some((_, static_init_ci, _, _)) = static_methods
            .iter()
            .find(|(mname, _, _, _)| mname.eq_ignore_ascii_case("__static_init__"))
        {
            let line = self.line;
            self.emit_u16(Op::REF_FUNC, *static_init_ci as u16);
            self.chunk().emit(0, line);
            self.emit_u8(Op::CALL_REF, 0);
            self.emit(Op::DROP);
        }

        // Inherit parent's static methods — walk up the chain via PendingClass
        if let Some(parent_name) = parent {
            let mut current_parent = Some(self.canon(parent_name));
            while let Some(ref pname) = current_parent {
                let parent_statics = self
                    .pending_classes
                    .get(pname.as_str())
                    .map(|pc| pc.statics.clone())
                    .unwrap_or_default();
                let next_parent = self
                    .pending_classes
                    .get(pname.as_str())
                    .and_then(|pc| pc.parent.clone());
                for (sname, sci) in &parent_statics {
                    // Only inherit if child doesn't already define it
                    if !all_statics.iter().any(|(n, _)| n == sname) {
                        common::classes::emit_attach_static_method(
                            self.chunk(),
                            ctor_local,
                            sname,
                            *sci,
                            php_static_receiver,
                            method_rest_fixed_count(*sci),
                            line,
                        );
                        all_statics.push((sname.clone(), *sci));
                    }
                }
                current_parent = next_parent;
            }
        }

        // Store statics in PendingClass for grandchildren to inherit
        if let Some(pc) = self.pending_classes.get_mut(name) {
            pc.statics = all_statics;
        }

        // Attach instance methods to the class object so static
        // `super.method()` dispatch can reach them. ECMA-262
        // §13.3.7.4 / §10.2.4 / §10.2.10.2: `super` resolves via
        // [[HomeObject]].[[Prototype]] (the parent class's prototype),
        // NOT the instance prototype chain. Multi-level inheritance
        // (C → B → A) needs B.method when called from C, A.method
        // when called from B — both at compile time. We mirror the
        // method bindings on the class constructor so
        // `GLOBAL_GET(ParentClass) ~ STRUCT_GET(method)` returns the
        // class-level method ref. Instance bindings are unchanged
        // (still per-instance for `this.method()` and override).
        for (mname, mci, _, _) in &instance_methods {
            // Skip getter/setter wrappers — they're bound differently.
            if mname.starts_with("__get_") || mname.starts_with("__set_") {
                continue;
            }
            common::classes::emit_attach_static_method(
                self.chunk(),
                ctor_local,
                mname,
                *mci,
                None,
                method_rest_fixed_count(*mci),
                line,
            );
        }

        let all_methods: Vec<(String, usize)> = method_chunks
            .iter()
            .map(|(n, c, _, _)| (n.clone(), *c))
            .collect();
        // Canonicalise per language case-sensitivity: case-insensitive
        // languages (VB/Pascal/COBOL/PHP) lowercase here, case-sensitive
        // (JS/TS/Python/C#) preserve. Registry stores whatever the walker
        // produced; runtime `Op::REF_TEST` looks up by the same canon.
        let canon_name = ctor_global_prefix;
        let canon_parent = parent.as_ref().map(|p| self.canon(p)).unwrap_or_default();
        common::classes::register_type(
            &mut self.chunks,
            &canon_name,
            &canon_parent,
            fields,
            all_methods,
            false,
            Vec::new(),
            Some(ctor_idx),
        );

        Ok(())
    }
}

fn body_has_result_member_assign(body: &[Statement]) -> bool {
    body.iter().any(stmt_has_result_member_assign)
}

fn stmt_has_result_member_assign(stmt: &Statement) -> bool {
    match &stmt.kind {
        StmtKind::Assign { targets, .. } => targets.iter().any(expr_is_result_member),
        StmtKind::Block(body)
        | StmtKind::FunctionDecl { body, .. }
        | StmtKind::With { body, .. }
        | StmtKind::Using { body, .. }
        | StmtKind::Lock { body, .. } => body_has_result_member_assign(body),
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            body_has_result_member_assign(then_body)
                || elifs
                    .iter()
                    .any(|(_, body)| body_has_result_member_assign(body))
                || else_body
                    .as_ref()
                    .is_some_and(|body| body_has_result_member_assign(body))
        }
        StmtKind::For { init, body, .. } => {
            init.as_ref()
                .is_some_and(|stmt| stmt_has_result_member_assign(stmt))
                || body_has_result_member_assign(body)
        }
        StmtKind::ForIn {
            body, else_body, ..
        }
        | StmtKind::While {
            body, else_body, ..
        } => {
            body_has_result_member_assign(body)
                || else_body
                    .as_ref()
                    .is_some_and(|body| body_has_result_member_assign(body))
        }
        StmtKind::DoWhile { body, .. } => body_has_result_member_assign(body),
        StmtKind::Switch { cases, default, .. } => {
            cases
                .iter()
                .any(|case| body_has_result_member_assign(&case.body))
                || default
                    .as_ref()
                    .is_some_and(|body| body_has_result_member_assign(body))
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
            ..
        } => {
            body_has_result_member_assign(body)
                || catches
                    .iter()
                    .any(|catch| body_has_result_member_assign(&catch.body))
                || else_body
                    .as_ref()
                    .is_some_and(|body| body_has_result_member_assign(body))
                || finally
                    .as_ref()
                    .is_some_and(|body| body_has_result_member_assign(body))
        }
        _ => false,
    }
}

fn expr_is_result_member(expr: &Expression) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Member { object, .. }
            if matches!(&object.kind, ExprKind::Ident(name) if name.eq_ignore_ascii_case("Result"))
    )
}
