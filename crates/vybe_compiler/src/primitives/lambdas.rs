//! Lambda compilation (explicit captures, direct) + component-call.
//!
//! Extracted from `primitives/calls.rs` (`impl Compiler`).

use super::*;

impl Compiler {
    pub(super) fn split_explicit_capture(capture: &str) -> (bool, &str) {
        if let Some(name) = capture.strip_prefix('&') {
            (true, name)
        } else {
            (false, capture)
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
        // A capture list names bindings in the ENCLOSING scope, so the strings
        // arrive spelled exactly as that scope defined them — sigil and all.
        // Nothing to normalize here: the walker is the only thing that knows
        // its own spelling, and it already applied it.
        if captures
            .iter()
            .any(|capture| !Self::split_explicit_capture(capture).0)
        {
            return self.compile_lambda_with_explicit_captures(
                params,
                body,
                captures,
                is_async,
                is_generator,
            );
        }

        self.compile_lambda_direct(params, body, is_async, is_generator, is_arrow)
    }

    pub(super) fn compile_lambda_with_explicit_captures(
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
        self.scopes.push(Scope::new_function(self.directives().variable_fold()));
        let saved = self.current;
        self.current = factory_idx;

        // The factory is a FUNCTION FRAME like any other, so it enters the
        // shared frame books. No seed: its only locals are the by-value
        // capture params, and the inner lambda's upvalues resolve through
        // them, never through an enclosing shared env.
        let frame_books = self.enter_closure_frame(&[]);

        for (capture_name, capture_type) in &capture_bindings {
            self.define_source_local_typed(capture_name, capture_type.clone().map(Into::into));
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
        self.exit_closure_frame(frame_books);

        let line = self.line;
        if uvs.is_empty() {
            common::functions::emit_ref_func(&mut self.chunks[self.current], factory_idx, 0, line);
        } else {
            let mut env_slots: Vec<u16> = Vec::new();
            for (i, uv) in uvs.iter().enumerate() {
                if let Some(name) = uv_names[i].clone() {
                    let slot = if uv.is_local {
                        uv.index
                    } else {
                        let parent_env = self.closure_env_slot();
                        let parent_idx = self.closure_env_index(&name);
                        let tmp = self.define_local(&format!("__nested_cap_{}", name));
                        crate::primitives::closures::emit_env_get(
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
            crate::primitives::closures::emit_env_new(self.chunk(), &env_slots, line);
            let env_slot = self.define_local(&format!("__closure_env_factory_{}", factory_idx));
            self.emit_u16(Op::LOCAL_SET, env_slot);
            common::functions::emit_ref_func(&mut self.chunks[self.current], factory_idx, 1, line);
            common::functions::emit_closure_upvalue(
                &mut self.chunks[self.current],
                true,
                env_slot,
                line,
            );
        }
        for capture in captures {
            let (by_ref, capture_name) = Self::split_explicit_capture(capture);
            if !by_ref {
                if self.ambient_this() && capture_name == "__js_this" {
                    self.compile_expr(&Expression::new(ExprKind::This))?;
                } else {
                    self.emit_var_get(capture_name);
                }
            }
        }
        self.emit_direct_callable_invoke(capture_bindings.len() as u8);
        Ok(())
    }

    pub(super) fn compile_lambda_direct(
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
        if self.ambient_this() && is_arrow {
            let self_kw = self.profile.self_keyword.clone();
            let scope_idx = self.scopes.len() - 1;
            let this_reachable = self.scope().resolve(&self_kw).is_some()
                || self.scope().resolve("__js_this").is_some()
                || (scope_idx > 0
                    && (self.resolve_upvalue(scope_idx, &self_kw).is_some()
                        || self.resolve_upvalue(scope_idx, "__js_this").is_some()));
            if !this_reachable {
                let slot = self.define_local("__js_this");
                self.emit_global_read("__js_this");
                self.emit_u16(Op::LOCAL_SET, slot);
            }
            let nt_reachable = self.scope().resolve("__js_new_target").is_some()
                || (scope_idx > 0 && self.resolve_upvalue(scope_idx, "__js_new_target").is_some());
            if !nt_reachable {
                let slot = self.define_local("__js_new_target");
                self.emit_global_read("__js_new_target");
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
        // Record the BODY chunk — not the factory that wraps it. A caller that
        // needs to stamp a fact on the lambda (an object-literal accessor
        // declaring its receiver convention) has no other handle: the chunk is
        // anonymous, and `compile_lambda_with_flags` hands back the FACTORY's
        // index, which is a different chunk entirely.
        self.last_lambda_body_chunk = Some(ci);
        self.chunks[ci].is_async = is_async;
        self.chunks[ci].is_generator = is_generator;
        // A closure resolves names the way the body containing it does — a PHP
        // closure sees no more of the module than the function it sits in.
        let enclosing = self.scope().resolution;
        self.scopes
            .push(Scope::new_function_like(enclosing, self.directives().variable_fold()));
        let saved = self.current;
        self.current = ci;
        // Runtime TRY_END counts are per-FRAME: a nested chunk must not
        // inherit the enclosing async body's try depth, or its returns pop the
        // CALLER's handlers off the shared runtime handler stack (a lambda
        // compiled inline inside an async fn emitted TRY_END × 2, silently
        // removing the user's enclosing try/catch).
        let saved_async_try_depth = std::mem::take(&mut self.active_async_try_depth);
        let saved_fn = self.current_func_name.replace("<lambda>".into());
        let frame_books = self.enter_closure_frame(&parent_shared_env_names);
        // ECMA-262 §11.2.2: strict mode is inherited by nested functions and
        // additionally turned on by a `"use strict"` directive prologue in
        // this function's own block body. Arrow expression bodies cannot carry
        // a prologue, so they only inherit.
        let saved_strict = self.in_strict;
        match body {
            LambdaBody::Block(stmts) => {
                if Self::stmts_have_use_strict_directive(stmts) {
                    self.in_strict = true;
                }
                crate::primitives::collect_closure_captured_idents(
                    stmts,
                    &mut self.current_closure_captured_locals,
                );
            }
            LambdaBody::Expr(expr) => {
                crate::primitives::collect_closure_captured_in_expr(
                    expr,
                    &mut self.current_closure_captured_locals,
                );
            }
        }
        for p in params {
            self.define_source_local_typed(&p.name, p.type_hint.clone());
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
        // A parameter is a declared binding like any other — narrow it here,
        // AFTER defaults, so it holds the value it actually ends up with.
        self.emit_param_type_bindings(params)?;
        // An alias parameter is handed a reference — mark it, do not wrap it.
        self.bind_alias_params(params);
        // Reserve the closure environment slot BEFORE any body statement is
        // compiled.
        //
        // `call_function_inner` copies upvalues into the frame at function
        // ENTRY, into `[capture_base .. capture_base + capture_count)`. But
        // `capture_local_slot` allocated that slot LAZILY, at the first
        // captured-variable access. Every statement compiled before that
        // access draws `alloc_scratch` slots out of the same low range —
        // and `compile_stmt` reclaims scratch back to the statement mark, so
        // the range genuinely is reused — overwriting the env the VM had
        // already placed there. The reclaim's `capture_locals` floor protects
        // the slot forward from the moment it exists; nothing protected it
        // backward, because the VM's write happens before statement one.
        //
        // A COBOL PROCEDURE DIVISION is exactly this shape: a capturing inner
        // lambda whose first statement is a DISPLAY, whose stream handles took
        // the env's slot. Reserving here is what makes `capture_base` true for
        // the whole body rather than only after the first read.
        //
        // Inert for a closure with no upvalues — the VM copies
        // min(upvalues.len(), capture_count) values, so a reserved-but-unused
        // slot costs one local and nothing else.
        self.closure_env_slot();
        // Snapshot __js_this as a local BEFORE shared env creation so inner
        // arrows can capture it via the shared env / upvalue chain.
        if self.ambient_this() && self.scopes.len() > 1 {
            let parent_has_this = self.scopes.len() > 2
                && self.scopes[self.scopes.len() - 2]
                    .resolve("__js_this")
                    .is_some();
            if !parent_has_this {
                let body_has_this = match body {
                    LambdaBody::Block(stmts) => crate::primitives::body_contains_this(stmts),
                    LambdaBody::Expr(expr) => crate::primitives::expr_contains_this(expr),
                };
                if body_has_this {
                    self.emit_global_read("__js_this");
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
                    self.emit_null();
                }
                self.chunks[self.current].emit_array_new_fixed(0, env_size, line);
                let env_slot = self.define_local("__shared_env");
                self.emit_u16(Op::LOCAL_SET, env_slot);
                self.shared_env_slot = Some(env_slot);
                self.shared_env_names = captured_names.clone();
                let mut local_decls: HashSet<String> =
                    params.iter().map(|p| p.name.clone()).collect();
                local_decls.insert("__js_this".to_string());
                if let LambdaBody::Block(stmts) = body {
                    crate::primitives::collect_declared_names(stmts, &mut local_decls);
                }
                for (idx, cap_name) in captured_names.iter().enumerate() {
                    if let Some(param_slot) = self.scope().resolve(cap_name) {
                        self.emit_u16(Op::LOCAL_GET, param_slot);
                        crate::primitives::closures::emit_env_set(
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
                            crate::primitives::closures::emit_env_get(
                                self.chunk(),
                                closure_env,
                                parent_idx as u16,
                                line,
                            );
                            crate::primitives::closures::emit_env_set(
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
            // The NAME comes from the profile, like the three sites in
            // `classes.rs` — hardcoding it captured any user variable spelled
            // the same in a case-insensitive language. Fortran writes `result`
            // as an ordinary identifier constantly; Pascal really does spell
            // the slot, which is why the default stays `Result`.
            let slot_name = self.profile.result_slot_name.clone();
            let rs = self.define_local(&slot_name);
            self.emit_null();
            self.emit_u16(Op::LOCAL_SET, rs);
            let saved_rs = self.current_result_slot.take();
            self.current_result_slot = Some(rs);
            Some((rs, saved_rs))
        } else {
            None
        };
        let saved_result_slot = result_slot.as_ref().map(|(_, saved_rs)| *saved_rs);

        // Same `!is_generator` gate as the declaration/method sites: an
        // async GENERATOR expression completes/throws through `resume`;
        // its promise surface is the attached `.next()` driver.
        let async_try = if is_async && !is_generator && self.profile.async_wraps_body_in_try {
            let line = self.line;
            { common::functions::emit_async_body_start(&mut self.chunks[self.current], line); Some(()) }
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

        if async_try.is_some() {
            let line = self.line;
            let chunk = &mut self.chunks[self.current];
            common::functions::emit_async_body_fallthrough(chunk, line);
            let resolve_idx = self.import("ecma:promise", "resolve");
            self.emit_host_call(resolve_idx, 1);
            self.emit(Op::RETURN);
            let chunk = &mut self.chunks[self.current];
            common::functions::end_async_body_handler(chunk, line);
            let reject_idx = self.import("ecma:promise", "reject");
            self.emit_host_call(reject_idx, 1);
            self.emit(Op::RETURN);
        } else if let Some((rs, saved_rs)) = result_slot {
            self.emit_u16(Op::LOCAL_GET, rs);
            self.emit(Op::RETURN);
            self.current_result_slot = saved_rs;
        } else if matches!(body, LambdaBody::Block(_)) {
            if self.profile.has_undefined_value {
                inst!(self, core_wasm::undefined);
                self.emit(Op::RETURN);
            } else {
                let line = self.line;
                common::functions::emit_function_epilogue(&mut self.chunks[ci], line);
            }
        }
        if let Some(saved_rs) = saved_result_slot {
            self.current_result_slot = saved_rs;
        }

        self.current_func_name = saved_fn;
        self.in_strict = saved_strict;

        let ns = self.scope().next_slot;
        self.chunks[ci].finalize_local_count(ns);
        self.chunks[ci].local_names = self.scope().defined_names.clone();
        let uvs = self.scopes.last().unwrap().upvalues.clone();
        // Resolve upvalue names BEFORE popping the inner scope
        let inner_scope_idx = self.scopes.len() - 1;
        let uv_names: Vec<Option<String>> = (0..uvs.len())
            .map(|i| self.captured_name_for_upvalue(inner_scope_idx, i as u8))
            .collect();
        self.scopes.pop();
        self.current = saved;
        self.active_async_try_depth = saved_async_try_depth;
        self.exit_closure_frame(frame_books);
        let parent_locals = self.scope().locals.clone();
        let line = self.line;
        if uvs.is_empty() {
            common::functions::emit_ref_func(&mut self.chunks[self.current], ci, 0, line);
        } else if let Some(shared_slot) = parent_shared_env_slot {
            // Parent has a shared env — pass it directly as the upvalue.
            // The inner function's closure_env_names was pre-seeded from
            // parent_shared_env_names, so indices match.
            common::functions::emit_ref_func(&mut self.chunks[self.current], ci, 1, line);
            common::functions::emit_closure_upvalue(
                &mut self.chunks[self.current],
                true,
                shared_slot,
                line,
            );
        } else {
            // No shared env — build a per-closure env (original path).
            let mut env_slots: Vec<u16> = Vec::new();
            let mut env_names: Vec<String> = Vec::new();
            for (i, uv) in uvs.iter().enumerate() {
                if let Some(name) = uv_names[i].clone() {
                    let slot = if uv.is_local {
                        let by_value = parent_locals
                            .iter()
                            .find(|l| l.slot == uv.index)
                            .map(|l| {
                                self.capture_by_value_vars
                                    .iter()
                                    .any(|n| *n == self.canon(&l.name))
                            })
                            .unwrap_or(false);
                        if by_value {
                            let orig_slot = uv.index;
                            self.emit_u16(Op::LOCAL_GET, orig_slot);
                            let snap = self.define_local(&format!("__snap_{}_{}", name, ci));
                            self.emit_u16(Op::LOCAL_SET, snap);
                            snap
                        } else {
                            uv.index
                        }
                    } else {
                        let parent_env = self.closure_env_slot();
                        let parent_idx = self.closure_env_index(&name);
                        let tmp = self.define_local(&format!("__nested_cap_{}", name));
                        crate::primitives::closures::emit_env_get(
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
            crate::primitives::closures::emit_env_new(self.chunk(), &env_slots, line);
            let env_slot = self.define_local(&format!("__closure_env_{}", ci));
            self.emit_u16(Op::LOCAL_SET, env_slot);
            common::functions::emit_ref_func(&mut self.chunks[self.current], ci, 1, line);
            common::functions::emit_closure_upvalue(
                &mut self.chunks[self.current],
                true,
                env_slot,
                line,
            );
        }
        if self.profile.has_function_prototype_bind {
            let length = params
                .iter()
                .take_while(|p| p.default.is_none() && !p.is_rest)
                .count();

            inst!(self, core_wasm::dup);
            self.emit_const(Value::F64(length as f64));
            let length_key = self.str_const("length");
            self.emit_struct_field_op(Op::STRUCT_SET, 0, length_key);

            inst!(self, core_wasm::dup);
            {
                // Recover the source kind when the walker lowered a generator
                // into a plain wrapper holding `__gen_fn` (obj-literal
                // `*m(){}` methods and generator expressions).
                let (eff_async, eff_generator) = match body {
                    LambdaBody::Block(stmts) => {
                        Self::wrapped_generator_kind(stmts).unwrap_or((is_async, is_generator))
                    }
                    _ => (is_async, is_generator),
                };
                let line = self.line;
                crate::primitives::prototypes::emit_stamp_function_kind_proto(
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
            self.emit_struct_field_op(Op::STRUCT_SET, 0, non_ctor_key);

            // §10.2.9/§10.2.10: name/length are non-enumerable.
            inst!(self, core_wasm::dup);
            {
                let line = self.line;
                crate::primitives::prototypes::emit_stamp_fn_metadata_nonenum(self.chunk(), line);
            }

            // §10.2.11: arrows carry a marker — the host uses it for
            // lexical-this (call/apply ignore thisArg) and toString's
            // `=>` form. Threaded from the ExprKind::Lambda compile arm;
            // object-literal shorthand methods pass false.
            if is_arrow {
                inst!(self, core_wasm::dup);
                self.emit_const(Value::Bool(true));
                let arrow_key = self.str_const("__fn_arrow");
                self.emit_struct_field_op(Op::STRUCT_SET, 0, arrow_key);
            }
        }
        if has_rest {
            self.emit_stamp_rest_metadata_on_stack(params.len().saturating_sub(1));
        }
        Ok(())
    }

    pub(super) fn try_compile_namespace_component_call(
        &mut self,
        parts: &[String],
        args: &[&Expression],
    ) -> Result<bool, String> {
        // This registers platform namespace trees as a side effect, then the
        // arity-aware type lookup below can see overloaded static members.
        let namespace_resolution = self.resolve_profile_namespace_chain(parts);

        if parts.len() >= 2 {
            let method_name = parts.last().expect("parts len checked");
            let class_name = parts[..parts.len() - 1].join(".");
            if let Some(member) = vybe_runtime::namespaces::lookup_type_static_member(
                &self.profile.namespaces.type_scopes,
                &class_name,
                method_name,
                self.tree_fold(),
            ) {
                if let Some(target) =
                    vybe_runtime::namespaces::select_overload(&member, args.len() as u8)
                {
                    match target {
                        vybe_runtime::namespaces::NamespaceNode::CommonEmit(emit) => {
                            if (emit.eq_ignore_ascii_case("dotnet.console_writeline")
                                || emit.eq_ignore_ascii_case("dotnet.console_write"))
                                && args.len() == 1
                            {
                                self.emit_dotnet_console_arg(args[0])?;
                            } else {
                                for a in args {
                                    self.compile_expr(a)?;
                                }
                            }
                            let line = self.line;
                            self.emit_common(emit, args.len() as u8, line);
                            return Ok(true);
                        }
                        vybe_runtime::namespaces::NamespaceNode::Fn {
                            module,
                            func,
                            arity,
                            ..
                        } => {
                            for a in args {
                                self.compile_expr(a)?;
                            }
                            let idx = self.import(module, func);
                            self.emit_host_call(idx, arity.unwrap_or(args.len() as u8));
                            return Ok(true);
                        }
                        _ => {}
                    }
                }
            }
        }

        // namespaceplan.md: platform surfaces are data in the shared tree;
        // the common resolver handles the mounted chain.
        match namespace_resolution {
            Some(super::resolver::Resolution::Tree(
                crate::primitives::namespaces::ResolutionTarget::CommonEmit(emit),
            )) => {
                if (emit.eq_ignore_ascii_case("dotnet.console_writeline")
                    || emit.eq_ignore_ascii_case("dotnet.console_write"))
                    && args.len() == 1
                {
                    self.emit_dotnet_console_arg(args[0])?;
                } else {
                    for a in args {
                        self.compile_expr(a)?;
                    }
                }
                let line = self.line;
                self.emit_common(&emit, args.len() as u8, line);
                Ok(true)
            }
            Some(
                super::resolver::Resolution::HostImport { module, func }
                | super::resolver::Resolution::Tree(
                    crate::primitives::namespaces::ResolutionTarget::HostCall {
                        module, func, ..
                    },
                ),
            ) => {
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
