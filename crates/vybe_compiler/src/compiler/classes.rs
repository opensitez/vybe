//! Class, constructor, and free-function compilation.
//!
//! Extracted from `compiler.rs` to keep that file navigable. The
//! methods on this `impl Compiler { ... }` block are private by
//! convention (they're only called from other compiler methods) and
//! crate-private for the `dotnet_register` bridge.

use super::*;
use crate::compiler::class_normalize::{BaseCall, NormalConstructor, NormalMethod};
use crate::compiler::ArrayBindingMetadata;

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
        } else if self.profile.has_undefined_value {
            // JS spec: declared fields with no initializer default to undefined (not null).
            inst!(self, core_wasm::undefined);
        } else {
            self.emit(Op::NULL);
        }
        self.emit_u16(Op::LOCAL_SET, value_slot);

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

    /// Instance identity + member stamps a derived constructor applies to
    /// the `this` produced by `super()`: __type / type-id, __super link,
    /// base-method saves, field initializers, instance-method binds, form
    /// identity and auto-init calls. Emitted right after a TOP-LEVEL
    /// `super()` statement, or (JS, when `super()` is nested — e.g.
    /// `try{super();}catch{}`) after the whole body guarded by
    /// `this != null`, since a nested super's completion point isn't
    /// statically known.
    #[allow(clippy::too_many_arguments)]
    /// Emit the parent constructor VALUE for class wiring. Bound parents
    /// (user classes, locals, globals) resolve normally; unbound intrinsic
    /// exception parents (`extends Error` / `extends RangeError` …) resolve
    /// through the canonical `__ctor_<Name>` anchors — the bare names are
    /// compile-time bindings, not vm globals, so a plain GLOBAL_GET yields
    /// null and silently breaks the prototype chain.
    fn emit_parent_ctor_value(&mut self, parent_name: &str) {
        let pname = self.canon(parent_name);
        let bound = self.scope().resolve(parent_name).is_some()
            || self.defined_globals.contains(&pname)
            || self.defined_classes.contains(&pname);
        if !bound
            && self.profile.ecma_error_object_shape
            && common::errors::is_exception_type(parent_name)
            && !self.shadows_builtin_type(parent_name)
        {
            let key = self.str_const(&format!("__ctor_{parent_name}"));
            self.emit_u16(Op::GLOBAL_GET, key);
        } else {
            self.emit_var_get(&pname);
        }
    }

    fn emit_derived_ctor_stamps(
        &mut self,
        name: &str,
        this_slot: u16,
        parent: &Option<String>,
        instance_method_names: &[String],
        field_inits: &[(
            String,
            Option<String>,
            Option<Expression>,
            Option<Vec<Expression>>,
        )],
        instance_methods: &[&(String, usize, bool, bool)],
        method_capture_name_map: &HashMap<usize, Vec<String>>,
        method_rest_fixed_counts: &HashMap<usize, u8>,
        is_value_type: bool,
        should_stamp_form_identity: bool,
        body_stmts: &[Statement],
        user_body: &[Statement],
        auto_init_methods: &[String],
        line: u32,
    ) -> Result<(), String> {
        self.emit_u16(Op::LOCAL_GET, this_slot);
        self.emit_const(Value::String(Arc::from(name)));
        let type_key = self.str_const("__type");
        self.emit_u16(Op::STRUCT_SET, type_key);
        self.emit(Op::DROP);
        let tid_key = self.str_const(&format!("__tid_{}", self.canon(name)));
        self.emit_u16(Op::LOCAL_GET, this_slot);
        self.emit_u16(Op::GLOBAL_GET, tid_key);
        self.emit(Op::SET_TYPE_ID);
        self.emit(Op::DROP);
        if let Some(parent_name) = parent {
            let pname = self.canon(parent_name);
            for method_name in instance_method_names {
                common::classes::emit_save_base_method(self.chunk(), this_slot, method_name, line);
            }
            self.emit_store_super_ref(this_slot, &pname);
        }
        for (fname, type_hint, init, array_bounds) in field_inits {
            self.emit_class_field_initializer(
                this_slot,
                fname,
                type_hint.as_deref(),
                init.as_ref(),
                array_bounds.as_deref(),
                is_value_type,
                line,
            )?;
        }
        for (mname, mci, _, _) in instance_methods {
            if mname.starts_with("__get_") {
                let prop = mname.strip_prefix("__get_").unwrap_or(mname);
                common::classes::emit_bind_getter(self.chunk(), this_slot, prop, *mci, line);
            } else if mname.starts_with("__set_") {
                let prop = mname.strip_prefix("__set_").unwrap_or(mname);
                common::classes::emit_bind_setter(self.chunk(), this_slot, prop, *mci, line);
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
                    method_rest_fixed_counts.get(mci).copied(),
                    !self.class_prototype_dispatch(),
                )?;
            }
        }
        if should_stamp_form_identity && !body_has_identity_stamp(body_stmts) {
            self.emit_form_identity_stamp(this_slot, name, line);
        }
        for aim in auto_init_methods {
            let has_method = instance_methods
                .iter()
                .any(|(n, _, _, _)| n.eq_ignore_ascii_case(aim));
            if has_method && !body_calls_method(user_body, aim) {
                common::classes::emit_auto_init_call(self.chunk(), this_slot, aim, line);
            }
        }
        Ok(())
    }

    pub(super) fn captured_name_for_upvalue(
        &self,
        scope_idx: usize,
        upvalue_idx: u8,
    ) -> Option<String> {
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
                .find(|local| local.slot == upvalue.index)
                .map(|local| local.name.clone())
        } else {
            self.captured_name_for_upvalue(parent_scope_idx, upvalue.index as u8)
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
                common::functions::emit_closure_upvalue(
                    &mut self.chunks[self.current],
                    true,
                    slot,
                    line,
                );
                continue;
            }
            if self.scopes.len() > 1 {
                if let Some(uv) = self.resolve_upvalue(self.scopes.len() - 1, capture_name) {
                    common::functions::emit_closure_upvalue(
                        &mut self.chunks[self.current],
                        false,
                        uv as u16,
                        line,
                    );
                    continue;
                }
            }
            return Err(format!(
                "failed to resolve captured class method binding '{capture_name}'"
            ));
        }
        Ok(())
    }

    /// Set `__js_new_target` to this class ONLY when currently unset.
    /// §13.3.7.1: `super()` preserves the active new.target — the chain's
    /// outermost `new` already set it; this default covers constructors
    /// invoked without an outer `new` frame.
    fn emit_default_js_new_target(&mut self, name: &str) {
        if !self.profile.ecma_new_dispatch {
            return;
        }
        let nt_key = self.str_const("__js_new_target");
        self.emit_u16(Op::GLOBAL_GET, nt_key);
        self.emit(Op::REF_IS_NULL);
        let line = self.line;
        self.chunks[self.current].emit_if(line);
        self.emit_var_get(name);
        self.emit_u16(Op::GLOBAL_SET, nt_key);
        self.chunks[self.current].emit_end(line);
    }

    /// Push the prototype object for a new instance's `__proto__` link:
    /// `(__js_new_target ?? <OwnClass>).prototype`. §9.1.13
    /// OrdinaryCreateFromConstructor uses new.target's prototype, so
    /// `this.constructor` resolves to the *invoked* class even while a
    /// parent constructor body runs under `super()`.
    fn emit_load_instance_proto(&mut self, class_name: &str) {
        let nt_key = self.str_const("__js_new_target");
        self.emit_u16(Op::GLOBAL_GET, nt_key);
        inst!(self, core_wasm::dup);
        self.emit(Op::REF_IS_NULL);
        let line = self.line;
        self.chunks[self.current].emit_if(line);
        self.emit(Op::DROP);
        let class_global = self.str_const(class_name);
        self.emit_u16(Op::GLOBAL_GET, class_global);
        self.chunks[self.current].emit_end(line);
        let prototype_key = self.str_const("prototype");
        self.emit_u16(Op::STRUCT_GET, prototype_key);
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

        // Prototype-dispatch profiles: the prototype is the source of truth,
        // so reassignment (`C.prototype.m = wrap(C.prototype.m)`) reaches
        // instances constructed afterwards. Falls back to the compiled ref
        // when the prototype has no entry (capture-carrying methods, class
        // expressions without a class global).
        let proto_class = if self.class_prototype_dispatch() && capture_names.is_empty() {
            self.current_class
                .clone()
                .filter(|c| self.defined_classes.contains(&self.canon(c)))
        } else {
            None
        };

        let mut bind_names = vec![method_name.to_string()];
        for &alias in common::classes::cross_language_aliases(method_name) {
            if alias != method_name {
                bind_names.push(alias.to_string());
            }
        }

        for bind_name in bind_names {
            self.emit_u16(Op::LOCAL_GET, this_slot);
            if let Some(class_name) = &proto_class {
                let cname = self.canon(class_name);
                let class_idx = self.global_name_const_idx(&cname);
                self.emit_u16(Op::GLOBAL_GET, class_idx);
                let proto_key = self.str_const("prototype");
                self.emit_u16(Op::STRUCT_GET, proto_key);
                let mkey = self.str_const(method_name);
                self.emit_u16(Op::STRUCT_GET, mkey);
                inst!(self, core_wasm::dup);
                self.emit(Op::REF_IS_NULL);
                let line = self.line;
                self.chunk().emit_if(line);
                self.emit(Op::DROP);
                self.emit_ref_func_with_captures(method_chunk_idx, capture_names)?;
                self.chunks[self.current].emit_end(line);
            } else {
                self.emit_ref_func_with_captures(method_chunk_idx, capture_names)?;
            }
            if bind_receiver {
                inst!(self, core_wasm::dup);
                self.emit_u16(Op::LOCAL_GET, this_slot);
                self.emit_u16(Op::STRUCT_SET, receiver_key);
                self.emit(Op::DROP);
            }
            if let Some(fixed_count) = rest_fixed_count {
                inst!(self, core_wasm::dup);
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

    /// Detect the JS walker's `wrap_generator` lowering: a plain outer
    /// function whose body binds `__gen_fn` to a generator function
    /// expression. Returns the SOURCE function's `(is_async, is_generator)`
    /// so prototype stamping reflects the original kind (§27.3/§27.4).
    pub(super) fn wrapped_generator_kind(body: &[Statement]) -> Option<(bool, bool)> {
        for stmt in body {
            if let StmtKind::VarDecl { declarations, .. } = &stmt.kind {
                for d in declarations {
                    if matches!(&d.pattern, crate::ast::BindingPattern::Ident(n) if n == "__gen_fn")
                    {
                        if let Some(init) = &d.init {
                            if let crate::ast::ExprKind::FunctionExpr(inner) = &init.kind {
                                if let StmtKind::FunctionDecl {
                                    is_async,
                                    is_generator,
                                    ..
                                } = &inner.kind
                                {
                                    if *is_generator {
                                        return Some((*is_async, true));
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        None
    }

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

        let uses_js_arguments = self.profile.has_arguments_object
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
        // `is_async` carries SOURCE truth (async fns AND async generators).
        // Consumers refine: the JSPI custom-section writer and the VM's
        // call_async gate both exclude generators (async generators are
        // continuations at call time — their async surface is `.next()`
        // returning a promise, which the protocol attach selects on this
        // flag).
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
        if self.profile.name == "php" {
            self.php_function_globals.push(HashSet::new());
        }
        let saved = self.current;
        self.current = func_idx;
        // Runtime TRY_END counts are per-FRAME: a nested chunk must not
        // inherit the enclosing async body's try depth, or its returns pop
        // the caller's handlers off the shared runtime handler stack.
        let saved_async_try_depth = std::mem::take(&mut self.active_async_try_depth);
        // Function body opens fresh wrt the runtime label_stack —
        // emit_return drains back to this base. Save+restore so nested
        // function decls compose.
        let saved_label_base = self.function_label_base;
        self.function_label_base = self.label_depth;
        let saved_fn = self.current_func_name.take();
        self.current_func_name = Some(name.to_string());
        // Pre-scan: collect locals/params whose address is taken (`&v`) so they
        // can be promoted to a pointer cell once at declaration / entry, rather
        // than lazily at the `&v` site (which re-wraps every loop iteration).
        // Scoped to C — the only profile exercising this path — so the other
        // AddrOf-using languages (Pascal/Go/C#) keep their current behavior.
        let saved_addr_taken = std::mem::take(&mut self.current_addr_taken_locals);
        if self.profile.name == "c" {
            crate::compiler::collect_addr_taken_idents(body, &mut self.current_addr_taken_locals);
        }
        let saved_closure_captured = std::mem::take(&mut self.current_closure_captured_locals);
        let saved_env_names = std::mem::take(&mut self.closure_env_names);
        let saved_capture_locals = std::mem::take(&mut self.capture_locals);
        // Capture parent shared env for nested function upvalue resolution
        let parent_shared_env_slot = self.shared_env_slot;
        let parent_shared_env_names = self.shared_env_names.clone();
        let saved_shared_env_slot = self.shared_env_slot.take();
        let saved_shared_env_names = std::mem::take(&mut self.shared_env_names);
        crate::compiler::collect_closure_captured_idents(
            body,
            &mut self.current_closure_captured_locals,
        );
        // If parent has a shared env, pre-seed closure_env_names so
        // upvalue indices match the parent's shared env layout.
        if !parent_shared_env_names.is_empty() {
            self.closure_env_names = parent_shared_env_names.clone();
        }
        // ECMA-262 §11.2.2: inherit strict mode and additionally enable it on
        // a `"use strict"` directive prologue in this function's body.
        let saved_strict = self.in_strict;
        if Self::stmts_have_use_strict_directive(body) {
            self.in_strict = true;
        }
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
            self.emit_u16(Op::LOCAL_GET, slot);
            self.emit_var_get(name);
            let callee_key = self.str_const("callee");
            self.emit_u16(Op::STRUCT_SET, callee_key);
            self.emit(Op::DROP);
            // §10.4.4.6: arguments objects report "[object Arguments]" —
            // stamp the tag the host's object_to_string_tag reads.
            self.emit_u16(Op::LOCAL_GET, slot);
            self.chunk().emit_string_const("Arguments", 0);
            let type_key = self.str_const("__type");
            self.emit_u16(Op::STRUCT_SET, type_key);
            self.emit(Op::DROP);
            Some(slot)
        } else {
            None
        };

        let mut aliased_params = HashMap::new();
        let mut aliased_indices = HashMap::new();
        // ECMA-262 §10.4.4: only a *non-strict* function with a simple
        // parameter list gets a mapped `arguments` object whose elements
        // alias the named parameters. Strict functions get an unmapped
        // (independent) copy, so `arguments[0] = …` must NOT change the
        // parameter (and vice versa).
        let simple_arguments_alias = uses_js_arguments
            && !self.in_strict
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
                } else {
                    self.emit_array_value_or_undefined(
                        js_arguments_source_slot.unwrap(),
                        js_arguments_len_slot.unwrap(),
                        index,
                    );
                    self.emit_u16(Op::LOCAL_SET, slot);
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
                if self.profile.missing_arg_is_undefined {
                    fn_call!(self, "wasm:js-undefined", "test", 1);
                } else {
                    self.emit(Op::REF_IS_NULL);
                }
                let branch_line = self.line;
                self.chunks[self.current].emit_if(branch_line);
                self.compile_expr(default)?;
                self.emit_u16(Op::LOCAL_SET, slot);
                self.chunks[self.current].emit_end(branch_line);
            }
            self.maybe_initialize_fortran_out_param(p);
        }

        // Promote address-taken params to a pointer cell once, at entry. A later
        // `&param` (e.g. inside a loop) then reuses this cell instead of
        // re-wrapping it each iteration. Reads/writes of the param are already
        // cell-aware once the binding is marked.
        if !self.current_addr_taken_locals.is_empty() {
            for p in params {
                if self.current_addr_taken_locals.contains(&p.name) {
                    self.promote_local_binding_to_pointer_cell(&p.name);
                }
            }
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
                let rs = self.define_local_typed(&slot_name, return_type.clone());
                self.emit(Op::NULL);
                self.emit_u16(Op::LOCAL_SET, rs);
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
        let async_try = if is_async && !is_generator && self.profile.async_wraps_body_in_try {
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

        if self.profile.ambient_this_binding
            && crate::compiler::closures_in_body_reference_this(body)
        {
            let this_idx = self.str_const("__js_this");
            self.emit_u16(Op::GLOBAL_GET, this_idx);
            let this_local = self.define_local("__js_this");
            self.emit_u16(Op::LOCAL_SET, this_local);
            self.current_closure_captured_locals
                .insert("__js_this".to_string());
        }

        if !self.current_closure_captured_locals.is_empty() {
            let mut fn_scope_names: HashSet<String> =
                params.iter().map(|p| p.name.clone()).collect();
            crate::compiler::collect_declared_names(body, &mut fn_scope_names);
            let mut captured_names: Vec<String> = self
                .current_closure_captured_locals
                .iter()
                .filter(|name| {
                    fn_scope_names.contains(name.as_str())
                        || parent_shared_env_names.iter().any(|n| n == name.as_str())
                })
                .cloned()
                .collect();
            captured_names.sort();
            if !captured_names.is_empty() {
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
                crate::compiler::collect_declared_names(body, &mut local_decls);
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
        self.current_addr_taken_locals = saved_addr_taken;
        self.current_closure_captured_locals = saved_closure_captured;
        self.closure_env_names = saved_env_names;
        self.capture_locals = saved_capture_locals;
        self.shared_env_slot = saved_shared_env_slot;
        self.shared_env_names = saved_shared_env_names;
        self.in_strict = saved_strict;
        self.current_result_slot = saved_rs;
        self.current_ref_out_params = saved_ref_out;

        let ns = self.scope().next_slot;
        self.chunks[func_idx].finalize_local_count(ns);
        let uvs = self.scopes.last().unwrap().upvalues.clone();
        let inner_scope_idx = self.scopes.len() - 1;
        let uv_names: Vec<Option<String>> = (0..uvs.len())
            .map(|i| self.captured_name_for_upvalue(inner_scope_idx, i as u8))
            .collect();
        self.js_arguments_bindings.pop();
        self.scopes.pop();
        self.static_local_bindings.pop();
        if self.profile.name == "php" {
            self.php_function_globals.pop();
        }
        self.current = saved;
        self.active_async_try_depth = saved_async_try_depth;
        self.function_label_base = saved_label_base;

        let line = self.line;
        if uvs.is_empty() {
            common::functions::emit_ref_func(&mut self.chunks[self.current], func_idx, 0, line);
        } else if let Some(shared_slot) = parent_shared_env_slot {
            // Parent has a shared env — pass it directly as the upvalue.
            common::functions::emit_ref_func(&mut self.chunks[self.current], func_idx, 1, line);
            common::functions::emit_closure_upvalue(
                &mut self.chunks[self.current],
                true,
                shared_slot,
                line,
            );
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
            let env_slot = self.define_local(&format!("__closure_env_{}", func_idx));
            self.emit_u16(Op::LOCAL_SET, env_slot);
            common::functions::emit_ref_func(&mut self.chunks[self.current], func_idx, 1, line);
            common::functions::emit_closure_upvalue(
                &mut self.chunks[self.current],
                true,
                env_slot,
                line,
            );
        }
        if uses_js_arguments {
            self.emit_stamp_rest_metadata_on_stack(0);
        } else if has_rest {
            self.emit_stamp_rest_metadata_on_stack(params.len().saturating_sub(1));
        }
        let idx = self.str_const(name);
        self.emit_u16(Op::GLOBAL_SET, idx);
        if self.is_php_profile() {
            self.emit_var_get(name);
            let php_fn_idx = self.str_const(&format!("__php_func${}", name));
            self.emit_u16(Op::GLOBAL_SET, php_fn_idx);
        }

        if self.profile.has_function_prototype_bind {
            let line = self.line;
            self.emit_common("object.new", 0, line);
            let proto_slot = self.define_local("__js_fn_proto");
            self.emit_u16(Op::LOCAL_SET, proto_slot);

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

            // The JS walker's wrap_generator lowers `function*` /
            // `async function*` to a PLAIN outer function holding
            // `const __gen_fn = function*(){...}` — recover the source
            // kind from that contract so the §27.3/§27.4 intrinsic
            // stamp survives the lowering.
            let (eff_async, eff_generator) =
                Self::wrapped_generator_kind(body).unwrap_or((is_async, is_generator));
            self.emit_var_get(name);
            {
                let line = self.line;
                crate::emitter::prototypes::emit_stamp_function_kind_proto(
                    self.chunk(),
                    eff_async,
                    eff_generator,
                    line,
                );
            }

            // §10.2.9/§10.2.10: name/length are non-enumerable.
            self.emit_var_get(name);
            {
                let line = self.line;
                crate::emitter::prototypes::emit_stamp_fn_metadata_nonenum(self.chunk(), line);
            }

            // §27.3 / §27.7: generator and async function declarations
            // have no [[Construct]] — `new` on them must TypeError (the
            // host construct path checks this marker).
            if eff_async || eff_generator {
                self.emit_var_get(name);
                self.emit_const(Value::Bool(true));
                let non_ctor_key = self.str_const("__vybe_non_ctor");
                self.emit_u16(Op::STRUCT_SET, non_ctor_key);
                self.emit(Op::DROP);
            }

            // §27.7 / §27.3 (node-verified): async (non-generator)
            // functions have NO own `prototype` property; generator
            // functions have one WITHOUT a `constructor` property; plain
            // functions get the classic prototype/constructor pair.
            if !eff_async || eff_generator {
                if !eff_generator {
                    self.emit_u16(Op::LOCAL_GET, proto_slot);
                    self.emit_var_get(name);
                    let ctor_key = self.str_const("constructor");
                    self.emit_u16(Op::STRUCT_SET, ctor_key);
                    self.emit(Op::DROP);
                }

                self.emit_var_get(name);
                self.emit_u16(Op::LOCAL_GET, proto_slot);
                let proto_key = self.str_const("prototype");
                self.emit_u16(Op::STRUCT_SET, proto_key);
                self.emit(Op::DROP);
            }
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

    /// Recursively register a minimal member surface for every nested
    /// class/struct in `members`, keyed by its (already-qualified) name, so
    /// a reference to it from a sibling method compiled earlier resolves as
    /// a user type instead of falling through to a builtin value-method. The
    /// full registration replaces this when the nested type is compiled.
    fn predeclare_nested_type_surfaces(&mut self, members: &[ClassMember], enclosing: &str) {
        for m in members {
            let ClassMember::NestedType(stmt) = m else {
                continue;
            };
            let (nested_name, nested_members, nested_parent): (
                &str,
                &[ClassMember],
                Option<String>,
            ) = match &stmt.kind {
                StmtKind::ClassDecl {
                    name: nn,
                    members: nm,
                    parents,
                    ..
                } => (nn, nm, parents.first().map(|p| self.canon(p))),
                StmtKind::StructDecl {
                    name: nn,
                    members: nm,
                    ..
                } => (nn, nm, None),
                _ => continue,
            };
            let nested_canon = self.canon(nested_name);
            if !self.pending_classes.contains_key(&nested_canon) {
                let mut static_method_names: Vec<String> = Vec::new();
                let mut static_fields: Vec<String> = Vec::new();
                let mut instance_member_names: Vec<String> = Vec::new();
                let mut fields: Vec<String> = Vec::new();
                let mut field_storage_names: HashMap<String, String> = HashMap::new();
                let field_storage_slot_name =
                    |compiler: &Self, owner_class: &str, field_name: &str| {
                        let field_canon = compiler.canon(field_name);
                        if compiler.profile.field_hiding
                            && compiler.field_hides_ancestor(nested_parent.as_deref(), &field_canon)
                        {
                            format!("__hide_{}${}", compiler.canon(owner_class), field_canon)
                        } else {
                            compiler.js_member_storage_name_for_class(owner_class, field_name)
                        }
                    };
                // (method-return-type key, return type) — registered so a
                // chained call `outer.first().next()` compiled in a sibling
                // method (before this nested type is compiled) can infer the
                // intermediate result's type.
                let mut return_types: Vec<(String, String)> = Vec::new();
                for mem in nested_members {
                    match mem {
                        ClassMember::Method(ms) => {
                            if let StmtKind::FunctionDecl {
                                name: mname,
                                modifiers,
                                return_type,
                                ..
                            } = &ms.kind
                            {
                                if modifiers.is_abstract {
                                    continue;
                                }
                                if modifiers.is_static || modifiers.is_shared {
                                    static_method_names.push(self.canon(mname));
                                } else {
                                    instance_member_names.push(self.canon(mname));
                                }
                                if let Some(rt) = return_type {
                                    return_types.push((
                                        self.canon(&format!("{nested_canon}.{mname}")),
                                        rt.clone(),
                                    ));
                                }
                            }
                        }
                        ClassMember::Property { name: pname, .. } => {
                            instance_member_names.push(self.canon(pname));
                        }
                        ClassMember::Field {
                            name: fname,
                            modifiers,
                            ..
                        } => {
                            if modifiers.is_static || modifiers.is_shared {
                                static_fields.push(self.canon(fname));
                            } else {
                                let field_canon = self.canon(fname);
                                let storage_name =
                                    field_storage_slot_name(self, &nested_canon, fname);
                                if storage_name != field_canon {
                                    field_storage_names
                                        .insert(field_canon.clone(), storage_name.clone());
                                }
                                fields.push(storage_name);
                            }
                        }
                        _ => {}
                    }
                }
                self.defined_globals.insert(nested_canon.clone());
                self.defined_classes.insert(nested_canon.clone());
                self.note_pending_class(&nested_canon, nested_parent);
                if let Some(pc) = self.pending_classes.get_mut(&nested_canon) {
                    pc.enclosing_class = Some(enclosing.to_string());
                    pc.static_method_names = static_method_names;
                    pc.static_fields = static_fields;
                    pc.instance_member_names = instance_member_names;
                    pc.fields = fields;
                    pc.field_storage_names = field_storage_names;
                }
                for (key, rt) in return_types {
                    self.function_return_types.entry(key).or_insert(rt);
                }
            }
            // Recurse: register deeper nested types (`Outer.Inner.Deep`).
            self.predeclare_nested_type_surfaces(nested_members, &nested_canon);
        }
    }

    pub(crate) fn compile_class(
        &mut self,
        class: &crate::compiler::class_normalize::NormalClass,
    ) -> Result<(), String> {
        // Extract the canonicalised names the orchestration below needs.
        // Canonicalisation happens once here rather than at every caller.
        let cname = self.canon(&class.name);
        let name: &str = &cname;
        let parent_canonical = class.parent.as_ref().map(|p| self.canon(p));
        let parent: &Option<String> = &parent_canonical;

        // The ENCLOSING frame's shared env, captured before any method compile
        // clears it. A class declared inside a function closes over that frame,
        // and a captured local there does not live in a slot — it lives in the
        // env array, which every closure receives as upvalue[0]
        // (`bindings.rs::capture_local_slot`). So a method reading such a local
        // emits `env[idx]`, and its binding must therefore capture the ENV, not
        // the individual name. Empty for a top-level class, which is the gate:
        // no env, nothing to forward.
        let enclosing_shared_env_names = self.shared_env_names.clone();

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
        // When properties and methods share the object namespace only in
        // some languages, a property whose name collides with a method needs
        // a distinct slot (see `separate_property_method_namespace`).
        let colliding_method_names: std::collections::HashSet<String> =
            if self.profile.separate_property_method_namespace {
                class
                    .instance_methods
                    .iter()
                    .map(|method| self.canon(&method.source_name))
                    .collect()
            } else {
                std::collections::HashSet::new()
            };
        let mut field_storage_names: HashMap<String, String> = HashMap::new();
        let field_storage_slot_name = |compiler: &Self,
                                       field_name: &str,
                                       method_names: &std::collections::HashSet<String>|
         -> String {
            let canon = compiler.canon(field_name);
            if compiler.profile.separate_property_method_namespace && method_names.contains(&canon)
            {
                format!("__prop${}", canon)
            } else {
                compiler.js_member_storage_name_for_class(&class.name, field_name)
            }
        };

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
            let field_canon = self.canon(&f.name);
            // Field hiding (java/C#/VB): a field that shadows an ancestor's
            // gets a declaring-class-qualified slot so both survive on the
            // object and access resolves by the reference's declared type.
            let fname = if self.profile.field_hiding
                && self.field_hides_ancestor(class.parent.as_deref(), &field_canon)
            {
                format!("__hide_{}${}", self.canon(&class.name), field_canon)
            } else {
                field_storage_slot_name(self, &f.name, &colliding_method_names)
            };
            if fname != field_canon {
                field_storage_names.insert(field_canon.clone(), fname.clone());
            }
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
                    field_storage_slot_name(self, auto_field_name, &colliding_method_names);
                if pname_canon != self.canon(auto_field_name) {
                    field_storage_names.insert(self.canon(auto_field_name), pname_canon.clone());
                }
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
                        field_storage_names
                            .get(&self.canon(&f.name))
                            .cloned()
                            .unwrap_or_else(|| {
                                self.js_member_storage_name_for_class(&class.name, &f.name)
                            }),
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
                field_storage_names,
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

        // Predeclare nested class/struct member surfaces before compiling
        // this class's methods. The real nested-type compilation happens
        // later (with the rest of `raw_extra_members`, after methods, so
        // chunk indices stay byte-identical), which means a call to a nested
        // class's method from a sibling method — e.g. `Inner.add(...)` or
        // `innerInstance.get()` inside `main` — would otherwise not see
        // `Inner` in `pending_classes` yet and fall through to a builtin
        // value-method of the same name (`get`, `add`, …). Top-level classes
        // avoid this via `predeclare_type_names`; nested classes (java/C#/VB)
        // had no equivalent. Recurses the whole nested tree so a deeply
        // qualified reference (`Outer.Inner`) is registered too. Each
        // placeholder is replaced by the full registration when the nested
        // type is actually compiled below.
        self.predeclare_nested_type_surfaces(&class.raw_extra_members, name);

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
            // An index operator makes `x[i]` on this type a method call. Record
            // it against the class so the index site can resolve it from the
            // receiver's static type instead of probing every index at runtime.
            if common::classes::cross_language_aliases(&m.source_name)
                .contains(&"__getitem__")
            {
                let cname = self.canon(&class.name);
                self.classes_with_indexer.insert(cname);
            }
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
            // The receiver (`this`) is bound ambiently from the call context
            // (`__js_this`) rather than passed as an explicit first positional
            // parameter. Capability-driven — not gated on the language name.
            let ambient_this = cc.profile.ambient_this_binding;
            let has_rest = user_params.last().map_or(false, |p| p.is_rest);
            let generator_control_arity = usize::from(m.is_generator && !has_rest);
            if has_rest {
                cc.rest_fixed_arities
                    .insert(user_params.len().saturating_sub(1) as u8);
            }
            // Whether this method carries an implicit leading receiver slot,
            // so its `arity` is `params + 1`. Mirrors the arity branches below.
            let has_receiver = if is_static_init {
                false
            } else if is_static {
                cc.profile.name == "php"
            } else if ambient_this {
                false
            } else {
                true
            };
            let arity = (user_params.len()
                + usize::from(has_receiver)
                + generator_control_arity) as u8;

            let ci = cc.chunks.len();
            let mut chunk = common::functions::create_function_chunk(mname, arity);
            chunk.is_method = has_receiver;
            chunk.param_count = user_params.len() as u8;
            // A WASM function's type shape (params → results) is what the
            // `call_indirect` runtime check compares. WASM functions have no
            // implicit receiver, so ALL declared params count — `user_params`
            // drops the phantom `self` that `explicit_self_param` assumes. The
            // result count is encoded as comma-joined placeholders in
            // `return_type` (None = a 0-result/void function, distinct from the
            // default 1-value ABI). Gated on the profile capability, not a name.
            if cc.profile.function_references {
                chunk.param_count = m.params.len() as u8;
                chunk.result_arity = m
                    .return_type
                    .as_ref()
                    .map(|rt| rt.split(',').count() as u8)
                    .unwrap_or(0);
            }
            chunk.is_async = m.is_async;
            chunk.is_generator = m.is_generator;
            // Source-level kind for prototype stamping at the attach sites —
            // survives walker lowerings that clear the outer flags. Generator
            // methods are lowered to a plain wrapper holding `__gen_fn`;
            // recover the source kind from that contract.
            let (src_async, src_gen) =
                Self::wrapped_generator_kind(&m.body).unwrap_or((m.is_async, m.is_generator));
            cc.method_fn_kinds.insert(ci, (src_async, src_gen));
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
            let saved_closure_captured = std::mem::take(&mut cc.current_closure_captured_locals);
            crate::compiler::collect_closure_captured_idents(
                &m.body,
                &mut cc.current_closure_captured_locals,
            );

            if !ambient_this && !is_static_init && (!is_static || cc.profile.name == "php") {
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
            if !ambient_this && !is_static_init && (!is_static || cc.profile.name == "php") {
                if class.explicit_self_param {
                    if let Some(self_param) = m.params.first() {
                        if self_param.name != self_kw {
                            let self_slot = cc.scope().resolve(&self_kw).unwrap();
                            let alias_slot = cc
                                .define_local_typed(&self_param.name, self_param.type_hint.clone());
                            cc.emit_u16(Op::LOCAL_GET, self_slot);
                            cc.emit_u16(Op::LOCAL_SET, alias_slot);
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
            // Same handling `compile_function_decl` gives a nested function: the
            // enclosing shared-env SLOT is only meaningful in the enclosing
            // frame, so clear it, and pre-seed `closure_env_names` with that
            // frame's layout so `env[idx]` reads here line up with the array
            // actually being passed.
            let saved_shared_env_slot = cc.shared_env_slot.take();
            let saved_shared_env_names = std::mem::take(&mut cc.shared_env_names);
            let saved_closure_env_names = std::mem::take(&mut cc.closure_env_names);
            if !enclosing_shared_env_names.is_empty() {
                cc.closure_env_names = enclosing_shared_env_names.clone();
            }
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
                    if cc.profile.missing_arg_is_undefined {
                        fn_call!(cc, "wasm:js-undefined", "test", 1);
                    } else {
                        cc.emit(Op::REF_IS_NULL);
                    }
                    let branch_line = cc.line;
                    cc.chunks[cc.current].emit_if(branch_line);
                    cc.compile_expr(default)?;
                    cc.emit_u16(Op::LOCAL_SET, slot);
                    cc.chunks[cc.current].emit_end(branch_line);
                }
            }

            if ambient_this
                && !is_static
                && crate::compiler::closures_in_body_reference_this(&m.body)
            {
                let this_idx = cc.str_const("__js_this");
                cc.emit_u16(Op::GLOBAL_GET, this_idx);
                let this_local = cc.define_local(&self_kw);
                cc.emit_u16(Op::LOCAL_SET, this_local);
                cc.current_closure_captured_locals.insert(self_kw.clone());
            }

            // Shared env for closures inside class methods: if the
            // method body has inner closures that capture the method's
            // locals, create a shared env array so mutations are visible
            // across all closures (same mechanism as compile_lambda_direct).
            if !cc.current_closure_captured_locals.is_empty() {
                let mut captured_names: Vec<String> = cc
                    .current_closure_captured_locals
                    .iter()
                    .filter(|name| !cc.defined_globals.contains(name.as_str()))
                    .cloned()
                    .collect();
                captured_names.sort();
                if !captured_names.is_empty() {
                    let env_size = captured_names.len() as u16;
                    let line = cc.line;
                    for _ in 0..env_size {
                        cc.emit(Op::NULL);
                    }
                    cc.chunks[cc.current].emit_op_u16(Op::ARRAY_NEW_FIXED, env_size, line);
                    let env_slot = cc.define_local("__shared_env");
                    cc.emit_u16(Op::LOCAL_SET, env_slot);
                    cc.shared_env_slot = Some(env_slot);
                    cc.shared_env_names = captured_names.clone();

                    let mut local_decls: std::collections::HashSet<String> =
                        user_params.iter().map(|p| p.name.clone()).collect();
                    if !ambient_this && !is_static_init && (!is_static || cc.profile.name == "php")
                    {
                        local_decls.insert(self_kw.clone());
                    }
                    crate::compiler::collect_declared_names(&m.body, &mut local_decls);

                    for (idx, cap_name) in captured_names.iter().enumerate() {
                        if let Some(param_slot) = cc.scope().resolve(cap_name) {
                            cc.emit_u16(Op::LOCAL_GET, param_slot);
                            crate::emitter::closures::emit_env_set(
                                cc.chunk(),
                                env_slot,
                                idx as u16,
                                line,
                            );
                        }
                    }
                }
            }

            let async_try = if m.is_async && !m.is_generator && cc.profile.async_wraps_body_in_try {
                let line = cc.line;
                Some(common::functions::emit_async_body_start(
                    &mut cc.chunks[ci],
                    line,
                ))
            } else {
                None
            };
            if async_try.is_some() {
                cc.active_async_try_depth += 1;
            }

            if is_ctor {
                // §15.7.14: class constructors require `new` (JS only —
                // other languages construct through their own paths).
                if cc.profile.ecma_new_dispatch {
                    let line = cc.line;
                    common::classes::emit_class_requires_new_guard(cc.chunk(), &class.name, line);
                }
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

            if async_try.is_some() {
                cc.active_async_try_depth = cc.active_async_try_depth.saturating_sub(1);
            }
            if let Some(catch_jump) = async_try {
                let line = cc.line;
                let chunk = &mut cc.chunks[ci];
                common::functions::emit_async_body_fallthrough(chunk, catch_jump, line);
                let resolve_idx = cc.import("ecma:promise", "resolve");
                cc.emit_host_call(resolve_idx, 1);
                cc.emit(Op::RETURN);
                let chunk = &mut cc.chunks[ci];
                common::functions::patch_async_body_catch(chunk, catch_jump);
                let reject_idx = cc.import("ecma:promise", "reject");
                cc.emit_host_call(reject_idx, 1);
                cc.emit(Op::RETURN);
            }

            cc.current_func_name = saved_fn;
            cc.current_result_slot = saved_rs;
            cc.current_ref_out_params = saved_ref_out;
            cc.current_member_is_static = saved_member_static;
            cc.current_closure_captured_locals = saved_closure_captured;
            cc.shared_env_slot = saved_shared_env_slot;
            cc.shared_env_names = saved_shared_env_names;
            cc.closure_env_names = saved_closure_env_names;

            let ns = cc.scope().next_slot;
            cc.chunks[ci].finalize_local_count(ns);
            let method_scope_idx = cc.scopes.len() - 1;
            let mut capture_names: Vec<String> = cc.scopes[method_scope_idx]
                .upvalues
                .iter()
                .enumerate()
                .filter_map(|(index, _)| {
                    cc.captured_name_for_upvalue(method_scope_idx, index as u8)
                })
                .collect();
            // `emit_var_get` registers an upvalue per NAME but emits the read as
            // `env[idx]` (bindings.rs), so these names describe what the body
            // reads, not what it receives: the body receives ONE upvalue, the
            // env array. Binding the names individually would hand the method a
            // raw value where it expects the array. Capture the env instead —
            // the same rule `compile_function_decl` applies when the parent has
            // a shared env ("pass it directly as the upvalue"). Only when every
            // capture lives in that env; anything else still binds by name.
            if !enclosing_shared_env_names.is_empty()
                && !capture_names.is_empty()
                && capture_names
                    .iter()
                    .all(|name| enclosing_shared_env_names.contains(name))
            {
                capture_names = vec!["__shared_env".to_string()];
            }
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
                // Virtuality is decided HERE, once, so the call path stays
                // language-agnostic. Keyword languages (C#/VB/Pascal) mark the
                // method; virtual-by-default languages (java/python/js/...)
                // carry no keyword and opt in via the profile instead. A
                // `static` or `is_not_overridable` member can never be
                // overridden, so it keeps its direct bind either way.
                let is_virtual = m.is_virtual
                    || m.is_override
                    || m.is_abstract
                    || (cc.profile.methods_virtual_by_default
                        && !is_static
                        && !m.raw_modifiers.is_not_overridable);
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
                        is_virtual,
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

                {
                    let ns = self.scope().next_slot;
                    self.chunks[ci].finalize_local_count(ns);
                }
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
                {
                    let ns = self.scope().next_slot;
                    self.chunks[ci].finalize_local_count(ns);
                }
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

        // Captures are re-resolved BY NAME in whichever frame the ref is being
        // emitted from — this helper is referenced from two different frames
        // (the constructor's arity dispatcher, and the class's defining scope).
        // An `UpvalueDesc`'s `(is_local, index)` are coordinates into ONE
        // parent frame, so replaying them verbatim in the other frame reads a
        // different slot entirely: a class declared inside a function had its
        // ctor helper capture the enclosing local by slot, then the dispatcher
        // replayed that slot against its own params and got `undefined`.
        // Re-resolving also threads the capture through the dispatcher's own
        // upvalue list, so the class closure is built carrying it.
        let emit_helper_ref = |cc: &mut Compiler,
                               helper_idx: usize,
                               helper_captures: &[String]|
         -> Result<(), String> {
            cc.emit_ref_func_with_captures(helper_idx, helper_captures)
        };

        let mut ctor_helpers: Vec<(usize, usize, usize, Vec<String>, Option<String>)> =
            Vec::new();
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
            // As for methods: the enclosing shared-env slot names a local in the
            // ENCLOSING frame, so it must not leak into this chunk — reading it
            // here would index this frame's slot instead. Cleared, with
            // `closure_env_names` pre-seeded to the enclosing layout so the
            // `env[idx]` reads match the array this ctor is handed.
            let saved_ctor_shared_env_slot = self.shared_env_slot.take();
            let saved_ctor_shared_env_names = std::mem::take(&mut self.shared_env_names);
            let saved_ctor_closure_env_names = std::mem::take(&mut self.closure_env_names);
            if !enclosing_shared_env_names.is_empty() {
                self.closure_env_names = enclosing_shared_env_names.clone();
            }
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
            // §15.7.14: class constructors require `new` (JS only).
            // `__js_new_target` is null on plain calls; every `new` chain
            // (incl. super()) sets or defaults it before this body runs.
            // Emitted AFTER param slots are claimed — emitter scratch
            // allocation before define_local shifts param slots (the
            // documented alloc_scratch/define_local collision).
            if self.profile.ecma_new_dispatch {
                let line = self.line;
                common::classes::emit_class_requires_new_guard(self.chunk(), name, line);
            }
            for (i, p) in user_params.iter().enumerate() {
                if let Some(Some(default)) = ctor_param_defaults.get(i) {
                    let slot = self.scope().resolve(p).unwrap();
                    self.emit_u16(Op::LOCAL_GET, slot);
                    if self.profile.missing_arg_is_undefined {
                        fn_call!(self, "wasm:js-undefined", "test", 1);
                    } else {
                        self.emit(Op::REF_IS_NULL);
                    }
                    // Result is already I32(0/1) — no dyn_to_bool needed
                    let branch_line = self.line;
                    self.chunks[self.current].emit_if(branch_line);
                    self.compile_expr(default)?;
                    self.emit_u16(Op::LOCAL_SET, slot);
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
            if self.profile.ambient_this_binding {
                let js_this = self.str_const("__js_this");
                self.emit_u16(Op::GLOBAL_GET, js_this);
                self.emit_u16(Op::LOCAL_SET, this_slot);
            }
            // §9.1.1.3.4 (JS): derived-constructor `this` TDZ context.
            // While this chunk's body compiles, `this` reads and `super()`
            // calls emit runtime guards against this_slot (null until
            // super() initializes it). Saved/restored so nested classes
            // compiled mid-body don't leak the context.
            let saved_derived_ctx = self.js_derived_ctor_ctx.take();
            if self.profile.ecma_new_dispatch && parent.is_some() && ctor_body.is_some() {
                self.js_derived_ctor_ctx = Some((self.current, this_slot));
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
                            !self.class_prototype_dispatch(),
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
                                    self.emit_default_js_new_target(name);
                                    self.emit_var_get(parent_name);
                                    for a in *bargs {
                                        self.compile_expr(a)?;
                                    }
                                    self.emit_u8(Op::CALL_REF, bargs.len() as u8);
                                    self.emit_u16(Op::LOCAL_SET, this_slot);
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
                                    self.emit_default_js_new_target(name);
                                    self.emit_var_get(parent_name);
                                    self.emit_u8(Op::CALL_REF, 0);
                                    self.emit_u16(Op::LOCAL_SET, this_slot);
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
                            self.emit_default_js_new_target(name);
                            self.emit_var_get(parent_name);
                            if synthesized_forward_args {
                                let parent_ctor_slot =
                                    self.define_local(&format!("__{}_parent_ctor", helper_name));
                                self.emit_u16(Op::LOCAL_SET, parent_ctor_slot);
                                let parent_called_slot =
                                    self.define_local(&format!("__{}_parent_called", helper_name));
                                inst!(self, core_wasm::i32_const, 0);
                                self.emit_u16(Op::LOCAL_SET, parent_called_slot);
                                for count in (1..=IMPLICIT_CTOR_FORWARD_ARGS).rev() {
                                    self.emit_u16(Op::LOCAL_GET, parent_called_slot);
                                    self.emit(Op::I32_EQZ);
                                    self.chunks[self.current].emit_if(line);
                                    self.emit_u16(Op::LOCAL_GET, (count - 1) as u16);
                                    self.emit(Op::REF_IS_NULL);
                                    self.emit(Op::I32_EQZ);
                                    self.chunks[self.current].emit_if(line);
                                    self.emit_u16(Op::LOCAL_GET, parent_ctor_slot);
                                    for arg_index in 0..count {
                                        self.emit_u16(Op::LOCAL_GET, arg_index as u16);
                                    }
                                    self.emit_u8(Op::CALL_REF, count);
                                    self.emit_u16(Op::LOCAL_SET, this_slot);
                                    inst!(self, core_wasm::i32_const, 1);
                                    self.emit_u16(Op::LOCAL_SET, parent_called_slot);
                                    self.chunks[self.current].emit_end(line);
                                    self.chunks[self.current].emit_end(line);
                                }
                                self.emit_u16(Op::LOCAL_GET, parent_called_slot);
                                self.emit(Op::I32_EQZ);
                                self.chunks[self.current].emit_if(line);
                                self.emit_u16(Op::LOCAL_GET, parent_ctor_slot);
                                self.emit_u8(Op::CALL_REF, 0);
                                self.emit_u16(Op::LOCAL_SET, this_slot);
                                self.chunks[self.current].emit_end(line);
                            } else {
                                for i in 0..user_arity {
                                    self.emit_u16(Op::LOCAL_GET, i as u16);
                                }
                                self.emit_u8(Op::CALL_REF, user_arity);
                                self.emit_u16(Op::LOCAL_SET, this_slot);
                            }
                        } else if self.profile.ecma_error_object_shape
                            && common::errors::is_exception_type(parent_name)
                            && !self.shadows_builtin_type(parent_name)
                        {
                            // §15.7.14 default derived ctor over an intrinsic
                            // error parent: super(message) — construct through
                            // the canonical exception shape so message/chain
                            // match a directly-constructed parent error.
                            self.emit_u16(Op::LOCAL_GET, 0);
                            self.emit(Op::REF_IS_NULL);
                            self.chunks[self.current].emit_if_value(line);
                            self.emit_const(Value::String(Arc::from("")));
                            self.chunks[self.current].emit_else(line);
                            self.emit_u16(Op::LOCAL_GET, 0);
                            self.chunks[self.current].emit_end(line);
                            self.emit_js_exception_ctor_from_message_value(parent_name)?;
                            self.emit_u16(Op::LOCAL_SET, this_slot);
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

                        if self.class_prototype_dispatch() {
                            let proto_link_key = self.str_const("__proto__");
                            let proto_local =
                                self.define_local(&format!("__{}_link_proto", helper_name));
                            self.emit_load_instance_proto(name);
                            self.emit_u16(Op::LOCAL_SET, proto_local);
                            self.emit_u16(Op::LOCAL_GET, proto_local);
                            self.emit(Op::REF_IS_NULL);
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
                                    !self.class_prototype_dispatch(),
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
                        let user_body = &body_stmts[preamble_end..];
                        // JS: when super() isn't a top-level statement
                        // (e.g. `try{super();}catch{}`), its completion
                        // point isn't statically known — defer the
                        // instance stamps until after the body, guarded by
                        // `this != null` (§9.1.1.3.4: this_slot stays null
                        // when super() never ran; the constructor-return
                        // TDZ guard throws the ReferenceError then).
                        let stamps_deferred = super_idx.is_none() && self.profile.ecma_new_dispatch;
                        if !stamps_deferred {
                            self.emit_derived_ctor_stamps(
                                name,
                                this_slot,
                                parent,
                                &instance_method_names,
                                &field_inits,
                                &instance_methods,
                                &method_capture_name_map,
                                &method_rest_fixed_counts,
                                class.is_value_type,
                                should_stamp_form_identity,
                                body_stmts,
                                user_body,
                                &class.auto_init_methods,
                                line,
                            )?;
                        }
                        for stmt in user_body {
                            self.compile_stmt(stmt)?;
                        }
                        if stamps_deferred {
                            self.emit_u16(Op::LOCAL_GET, this_slot);
                            self.emit(Op::REF_IS_NULL);
                            self.emit(Op::I32_EQZ);
                            self.chunks[self.current].emit_if(line);
                            self.emit_derived_ctor_stamps(
                                name,
                                this_slot,
                                parent,
                                &instance_method_names,
                                &field_inits,
                                &instance_methods,
                                &method_capture_name_map,
                                &method_rest_fixed_counts,
                                class.is_value_type,
                                should_stamp_form_identity,
                                body_stmts,
                                user_body,
                                &class.auto_init_methods,
                                line,
                            )?;
                            self.chunks[self.current].emit_end(line);
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
                                !self.class_prototype_dispatch(),
                            )?;
                        }
                    }
                    if self.class_prototype_dispatch() {
                        let proto_link_key = self.str_const("__proto__");
                        let proto_local =
                            self.define_local(&format!("__{}_link_proto_base", helper_name));
                        self.emit_load_instance_proto(name);
                        self.emit_u16(Op::LOCAL_SET, proto_local);
                        self.emit_u16(Op::LOCAL_GET, proto_local);
                        self.emit(Op::REF_IS_NULL);
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

                if self.is_php_profile() {
                    // PHP runtime class identity. This must run AFTER the
                    // ctor body / parent-ctor call — a child ctor receives
                    // `this` from the parent (synthesized forward OR
                    // `parent::__construct()` in the body) carrying the
                    // PARENT's type_id and constructor, so the child
                    // re-stamps:
                    //  - `constructor` → the class object, for `new static`
                    //    and get_class ($this.constructor.name); the JS
                    //    path gets this via the prototype chain instead.
                    //  - `__type` + WASM GC type_id, so instanceof
                    //    (REF_TEST fast path) sees the runtime class.
                    let class_global = self.str_const(name);
                    let ctor_key = self.str_const("constructor");
                    self.emit_u16(Op::LOCAL_GET, this_slot);
                    self.emit_u16(Op::GLOBAL_GET, class_global);
                    self.emit_u16(Op::STRUCT_SET, ctor_key);
                    self.emit(Op::DROP);
                    let canon_name = self.canon(name);
                    common::classes::emit_retype_object(self.chunk(), this_slot, &canon_name, line);
                }

                common::classes::emit_instanceof_chain(
                    &mut self.chunks,
                    self.current,
                    this_slot,
                    name,
                    line,
                );
                let mut interface_names = class.interfaces.clone();
                interface_names.extend(self.reflection_interfaces(name));
                let mut seen_interfaces = std::collections::HashSet::new();
                for interface_name in interface_names {
                    if !seen_interfaces.insert(self.canon(&interface_name)) {
                        continue;
                    }
                    common::classes::emit_instanceof_chain(
                        &mut self.chunks,
                        self.current,
                        this_slot,
                        &interface_name,
                        line,
                    );
                }
                // Set __proto__ link for prototype-dispatch classes.
                // Done just before return so this_slot is guaranteed valid.
                if self.class_prototype_dispatch() {
                    let proto_key = self.str_const("__proto__");
                    self.emit_load_instance_proto(name);
                    let tmp = self.define_local("__final_proto");
                    self.emit_u16(Op::LOCAL_SET, tmp);
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    self.emit(Op::REF_IS_NULL);
                    self.emit(Op::I32_EQZ);
                    self.chunks[self.current].emit_if(line);
                    self.emit_u16(Op::LOCAL_GET, this_slot);
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    self.emit_u16(Op::STRUCT_SET, proto_key);
                    self.emit(Op::DROP);
                    self.chunks[self.current].emit_end(line);
                }
                // §9.1.1.3.4 (JS): returning from a derived constructor
                // with `this` still uninitialized (super() missing, or its
                // throw was caught) is a ReferenceError.
                if self.js_derived_ctor_ctx == Some((self.current, this_slot)) {
                    common::classes::emit_this_initialized_guard(self.chunk(), this_slot, line);
                }
                common::classes::emit_constructor_return(self.chunk(), this_slot, line);
            }

            {
                let ns = self.scope().next_slot;
                self.chunks[helper_idx].finalize_local_count(ns);
            }
            // Names, not slot coordinates — see `emit_helper_ref`. Resolved
            // while the helper's scope is still on the stack, since
            // `captured_name_for_upvalue` walks it to name each upvalue.
            let helper_scope_idx = self.scopes.len() - 1;
            let mut helper_captures: Vec<String> = (0..self.scopes[helper_scope_idx].upvalues.len())
                .filter_map(|index| self.captured_name_for_upvalue(helper_scope_idx, index as u8))
                .collect();
            // Same rule the methods use: a ctor body reading an enclosing
            // captured local emits `env[idx]`, so it must receive the env array
            // rather than the individual values.
            if !enclosing_shared_env_names.is_empty()
                && !helper_captures.is_empty()
                && helper_captures
                    .iter()
                    .all(|name| enclosing_shared_env_names.contains(name))
            {
                helper_captures = vec!["__shared_env".to_string()];
            }
            self.scopes.pop();
            self.current = saved_cur;
            self.current_class = saved_class2;
            self.current_class_implicit_self = saved_implicit2;
            self.js_derived_ctor_ctx = saved_derived_ctx;
            self.shared_env_slot = saved_ctor_shared_env_slot;
            self.shared_env_names = saved_ctor_shared_env_names;
            self.closure_env_names = saved_ctor_closure_env_names;
            ctor_helpers.push((
                user_arity as usize,
                ctor_min_arity,
                helper_idx,
                helper_captures,
                ctor_variant.and_then(|c| c.named_name.clone()),
            ));
        }

        let ctor_idx = self.chunks.len();
        let ctor_arity = ctor_helpers
            .iter()
            .map(|(arity, _, _, _, _)| *arity)
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

        // Abstract class: mark for compile-time check in `new` expressions.
        if class.is_abstract {
            self.abstract_classes.insert(self.canon(name));
        }

        let js_ctor_relaxes_min_arity = self.profile.relaxed_call_arity;
        let helper_for_count = |count: usize| {
            ctor_helpers
                .iter()
                .filter(|(arity, min_arity, _, _, _)| {
                    count <= *arity && (js_ctor_relaxes_min_arity || count >= *min_arity)
                })
                .min_by_key(|(arity, _, _, _, _)| *arity)
        };
        for count in (1..=ctor_arity as usize).rev() {
            self.emit_u16(Op::LOCAL_GET, (count - 1) as u16);
            // Argument-presence test: a *missing* trailing arg is `undefined`,
            // so `new C(null)` must still dispatch to the 1-arg constructor
            // (`null` is a value, not an absent argument). Languages without a
            // distinct `undefined` fall back to the null test.
            if self.profile.missing_arg_is_undefined {
                fn_call!(self, "wasm:js-undefined", "test", 1);
            } else {
                self.emit(Op::REF_IS_NULL);
            }
            // Result is already I32(0/1) — no dyn_to_bool needed
            self.emit(Op::I32_EQZ);
            self.chunks[self.current].emit_if(line);
            if let Some((_, _, helper_idx, helper_captures, _)) = helper_for_count(count) {
                emit_helper_ref(self, *helper_idx, helper_captures)?;
                for arg_index in 0..count {
                    self.emit_u16(Op::LOCAL_GET, arg_index as u16);
                }
                self.emit_u8(Op::CALL_REF, count as u8);
                self.emit_return_through_finally(1)?;
            }
            self.chunks[self.current].emit_end(line);
        }
        if let Some((_, _, helper_idx, helper_captures, _)) = helper_for_count(0) {
            emit_helper_ref(self, *helper_idx, helper_captures)?;
            self.emit_u8(Op::CALL_REF, 0);
        } else {
            self.emit(Op::NULL);
        }
        self.emit_return_through_finally(1)?;
        {
            let ns = self.scope().next_slot;
            self.chunks[ctor_idx].finalize_local_count(ns);
        }
        let ctor_upvalues = self.scope().upvalues.clone();
        self.scopes.pop();
        self.current = saved_cur;

        let ctor_local = self.define_local(&format!("__{}_ctor", name));
        let uv_pairs: Vec<(bool, u16)> = ctor_upvalues
            .iter()
            .map(|uv| (uv.is_local, uv.index))
            .collect();
        let case_sensitive = self.profile.ecma_new_dispatch;
        common::classes::emit_store_constructor_with_upvalues(
            self.chunk(),
            name,
            ctor_idx,
            ctor_local,
            &uv_pairs,
            case_sensitive,
            line,
        );
        if self.is_php_profile() {
            // Stamp the declared class name on the ctor function so
            // `get_class($x)` ($x.constructor.name) returns it. The JS
            // branch below stamps `name` during prototype wiring; PHP
            // skips that block, so stamp here.
            self.emit_u16(Op::LOCAL_GET, ctor_local);
            self.emit_const(Value::String(Arc::from(name)));
            let name_key = self.str_const("name");
            self.emit_u16(Op::STRUCT_SET, name_key);
            self.emit(Op::DROP);
        }
        for (arity, _, helper_idx, helper_captures, named) in &ctor_helpers {
            emit_helper_ref(self, *helper_idx, helper_captures)?;
            let helper_global = format!("{}$arity{}", ctor_global_prefix, arity);
            let helper_idx_const = self.str_const(&helper_global);
            self.emit_u16(Op::GLOBAL_SET, helper_idx_const);
            // A named constructor (`Point.origin()`) is reached through the
            // class rather than by arity — several of them commonly share an
            // arity with each other and with the unnamed ctor. The helper
            // already allocates and returns the instance, so stamping it on
            // the class object makes `Point.origin(...)` an ordinary call,
            // the same shape a factory constructor already compiles to.
            if let Some(named) = named {
                self.emit_u16(Op::LOCAL_GET, ctor_local);
                emit_helper_ref(self, *helper_idx, helper_captures)?;
                let key = self.str_const(named);
                self.emit_u16(Op::STRUCT_SET, key);
                self.emit(Op::DROP);
            }
        }

        if self.class_prototype_dispatch() {
            self.emit_common("object.new", 0, line);
            let proto_local = self.define_local(&format!("__{}_prototype", name));
            self.emit_u16(Op::LOCAL_SET, proto_local);

            if let Some(parent_name) = parent {
                self.emit_parent_ctor_value(parent_name);
                let parent_proto_key = self.str_const("prototype");
                self.emit_u16(Op::STRUCT_GET, parent_proto_key);
                let parent_proto_local = self.define_local(&format!("__{}_parent_prototype", name));
                self.emit_u16(Op::LOCAL_SET, parent_proto_local);
                self.emit_u16(Op::LOCAL_GET, parent_proto_local);
                self.emit(Op::REF_IS_NULL);
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

            // §15.7.5 step 7 (JS): the class constructor's own
            // [[Prototype]] — the parent constructor for derived classes
            // (static inheritance walks it), %Function.prototype% for
            // base classes (C.bind / C.call / C.apply resolve through it).
            if self.class_prototype_dispatch() {
                if let Some(parent_name) = parent {
                    self.emit_u16(Op::LOCAL_GET, ctor_local);
                    self.emit_parent_ctor_value(parent_name);
                    let proto_link_key = self.str_const("__proto__");
                    self.emit_u16(Op::STRUCT_SET, proto_link_key);
                    self.emit(Op::DROP);
                } else {
                    self.emit_u16(Op::LOCAL_GET, ctor_local);
                    crate::emitter::prototypes::emit_stamp_function_kind_proto(
                        self.chunk(),
                        false,
                        false,
                        line,
                    );
                }
            }

            // The prototype is the class's open method table: every
            // instance method lands on it, so `C.prototype.m` resolves
            // and reassignment has a real target. Capture-carrying
            // methods are skipped — their upvalues bind in the
            // constructor's frame, not at definition scope.
            for (mname, mci, _, _) in &instance_methods {
                if mname.starts_with("__get_") || mname.starts_with("__set_") {
                    continue;
                }
                let has_captures = method_capture_name_map
                    .get(mci)
                    .is_some_and(|c| !c.is_empty());
                if has_captures {
                    continue;
                }
                let (m_async, m_gen) = self
                    .method_fn_kinds
                    .get(mci)
                    .copied()
                    .unwrap_or((self.chunks[*mci].is_async, self.chunks[*mci].is_generator));
                self.emit_u16(Op::LOCAL_GET, proto_local);
                self.emit_ref_func_with_captures(*mci, &[])?;
                // ECMA-262 function-kind stamp: async/generator methods'
                // __proto__ is the matching intrinsic prototype (§27.7.1 /
                // §27.3.1 / §27.4.1) — `getPrototypeOf(C.prototype.m)`.
                if m_async || m_gen {
                    inst!(self, core_wasm::dup);
                    let line = self.line;
                    crate::emitter::prototypes::emit_stamp_function_kind_proto(
                        self.chunk(),
                        m_async,
                        m_gen,
                        line,
                    );
                }
                // §10.2.9 SetFunctionName: a class method's `name` is its
                // property key (non-enumerable, like all fn metadata).
                inst!(self, core_wasm::dup);
                self.emit_const(Value::String(Arc::from(mname.as_str())));
                let name_key = self.str_const("name");
                self.emit_u16(Op::STRUCT_SET, name_key);
                {
                    let line = self.line;
                    crate::emitter::prototypes::emit_stamp_fn_metadata_nonenum(self.chunk(), line);
                }
                let key = self.str_const(mname);
                self.emit_u16(Op::STRUCT_SET, key);
                self.emit(Op::DROP);
            }
        }

        // Static field initializers run with the self-reference bound to the
        // class constructor object (ECMA-262 §15.7.10 — `this` inside a static
        // field initializer is the class itself), so `static y = this.x * 2`
        // can read sibling static fields. Bind the self-keyword to `ctor_local`
        // for the duration of the initializer emission.
        let static_self_kw = self.profile.self_keyword.clone();
        let static_self_slot = self.define_local(&static_self_kw);
        self.emit_u16(Op::LOCAL_GET, ctor_local);
        self.emit_u16(Op::LOCAL_SET, static_self_slot);

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
            let (m_async, m_gen) = self
                .method_fn_kinds
                .get(mci)
                .copied()
                .unwrap_or((self.chunks[*mci].is_async, self.chunks[*mci].is_generator));
            common::classes::emit_attach_static_method_kinded(
                self.chunk(),
                ctor_local,
                mname,
                *mci,
                php_static_receiver,
                method_rest_fixed_count(*mci),
                m_async,
                m_gen,
                line,
            );
            all_statics.push((mname.clone(), *mci));
        }

        if self.profile.supports_private_fields {
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

        // Attach instance methods/accessors to the class object so static
        // `super.method()` and `super.prop` dispatch can reach them. ECMA-262
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

        // Prototype-dispatch profiles resolve instance members through the
        // prototype chain — statics live on the constructor object only
        // (§15.7: `instance.staticMethod` is undefined), so keep them out
        // of the type table's instance-method fallback. Other dispatch
        // models (VB-style `Instance.SharedMethod`) keep the full list.
        let all_methods: Vec<(String, usize)> = method_chunks
            .iter()
            .filter(|(_, _, _, is_static)| !self.class_prototype_dispatch() || !*is_static)
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
            std::collections::HashMap::new(),
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
